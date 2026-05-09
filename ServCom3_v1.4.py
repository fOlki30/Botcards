#!/usr/bin/env python3
# card_bot3_screenshots_nodup.py
"""
Visit Steam market pages and capture screenshots:
- graph canvas
- sell orders table
- buy orders table
Outputs XLSX/ODS and saves progress.
Avoid duplicate output rows by building the output from the
input rows and the visited mapping (one line per URL).
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import logging
import os
import random
import tempfile
import time
import base64
from contextlib import contextmanager
from typing import List, Optional, Tuple

import pandas as pd
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.webdriver import (
    WebDriver as FirefoxDriver,
)
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm
from webdriver_manager.firefox import GeckoDriverManager

# constants
IN_FILE = "steam_games_cards.xlsx"
OUT_FILE = "steam_games_cards_images.xlsx"
URL_COL = "Market URL"
CARD_COL = "Nom de la carte"
GAME_COL = "Nom du jeu"
PAGE_TIMEOUT = 60

TEST_N = 20
DELAY_MIN = 7.0
DELAY_MAX = 10.0
BATCH = 50

EL_WAIT = 25
RETRY = 1
RETRY_BACK = 2.0

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("card_bot_screens_nodup")

UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/115.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_4) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) "
    "Version/16.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:115.0) "
    "Gecko/20100101 Firefox/115.0",
]


def build_driver(headless: bool = False, ua: Optional[str] = None) -> FirefoxDriver:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.set_preference("dom.webdriver.enabled", False)
    opts.set_preference("useAutomationExtension", False)
    if ua:
        opts.set_preference("general.useragent.override", ua)
    gecko = GeckoDriverManager().install()
    svc = Service(executable_path=gecko)
    drv = FirefoxDriver(service=svc, options=opts)
    drv.set_page_load_timeout(PAGE_TIMEOUT)
    return drv


@contextmanager
def firefox_driver(headless: bool = False, ua: Optional[str] = None):
    drv = build_driver(headless=headless, ua=ua)
    try:
        yield drv
    finally:
        try:
            drv.quit()
        except Exception:
            logger.exception("close driver error")


def read_input_rows(fn: str) -> List[Tuple[str, str, str]]:
    if not os.path.exists(fn):
        logger.error("file not found: %s", fn)
        return []
    try:
        df = pd.read_excel(fn, engine="openpyxl")
    except Exception:
        logger.exception("read excel failed: %s", fn)
        return []
    if GAME_COL in df.columns:
        games = df[GAME_COL].astype(str).tolist()
    elif df.shape[1] >= 1:
        games = df.iloc[:, 0].astype(str).tolist()
    else:
        logger.error("game col missing")
        return []
    if CARD_COL in df.columns:
        cards = df[CARD_COL].astype(str).tolist()
    elif df.shape[1] >= 2:
        cards = df.iloc[:, 1].astype(str).tolist()
    else:
        logger.error("card col missing")
        return []
    if URL_COL in df.columns:
        urls = df[URL_COL].astype(str).tolist()
    elif df.shape[1] >= 3:
        urls = df.iloc[:, 2].astype(str).tolist()
    else:
        logger.error("url col missing")
        return []
    rows: List[Tuple[str, str, str]] = []
    for i, u in enumerate(urls):
        g = games[i] if i < len(games) else ""
        c = cards[i] if i < len(cards) else ""
        rows.append((str(g).strip(), str(c).strip(), str(u).strip()))
    logger.info("rows read: %d", len(rows))
    return rows


def _choose_engine(pref: Optional[str] = None) -> str:
    if pref:
        return pref
    if importlib.util.find_spec("xlsxwriter") is not None:
        return "xlsxwriter"
    return "openpyxl"


def _atomic_write(path: str, data: str) -> None:
    fd, tmp = tempfile.mkstemp(dir=os.path.dirname(path) or ".")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(data)
        os.replace(tmp, path)
    except Exception:
        try:
            os.remove(tmp)
        except Exception:
            pass


def load_progress(pf: str) -> Tuple[dict, list]:
    if not os.path.exists(pf):
        return {}, []
    try:
        with open(pf, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        logger.exception("read progress failed")
        return {}, []
    visited = data.get("visited", {})
    results = [tuple(x) for x in data.get("results", [])]
    return visited, results


def save_progress(pf: str, visited: dict, results: list) -> None:
    # Keep saving both visited and results for compatibility,
    # but results may contain duplicates; visited is authoritative.
    payload = {"visited": visited, "results": results}
    try:
        _atomic_write(pf, json.dumps(payload, ensure_ascii=False))
    except Exception:
        logger.exception("save progress failed")


def save_output_from_input(
    input_rows: List[Tuple[str, str, str]],
    visited: dict,
    fn: str,
    engine: Optional[str] = None,
) -> None:
    """
    Build output rows from input_rows and visited mapping to avoid duplicates.
    For each input row (game, card, url) we write exactly one output line:
    - if url in visited -> use visited[url]
    - else -> blank image paths
    Duplicate URLs in input are written only once (first occurrence).
    """
    out_rows = []
    seen_urls = set()
    for game, card, url in input_rows:
        if not url or url.lower() in ("nan", "none"):
            # keep original empty entry
            out_rows.append((game, card, "", "", ""))
            continue
        if url in seen_urls:
            # skip duplicate input URL to avoid duplicate output lines
            continue
        seen_urls.add(url)
        v = visited.get(url)
        if v:
            # visited[url] is expected to be a list like [game, card, gpath, spath, bpath]
            # ensure length and convert to tuple
            try:
                row = tuple(v)
            except Exception:
                row = (game, card, "", "", "")
            out_rows.append(row)
        else:
            out_rows.append((game, card, "", "", ""))

    df = pd.DataFrame(
        out_rows,
        columns=[
            "Nom du jeu",
            "Nom de la carte",
            "Chemin image graphique",
            "Chemin image sell",
            "Chemin image buy",
        ],
    )
    base, ext = os.path.splitext(fn)
    ext = ext.lower()
    if ext not in (".xlsx", ".ods"):
        fn = base + ".xlsx"
        ext = ".xlsx"
    dirn = os.path.dirname(os.path.abspath(fn)) or "."
    fd, tmp = tempfile.mkstemp(suffix=ext, dir=dirn)
    os.close(fd)
    try:
        if ext == ".ods":
            try:
                df.to_excel(tmp, index=False, engine="odf")
                used = "odf"
            except Exception:
                logger.exception("ods write failed")
                eng = engine or _choose_engine()
                df.to_excel(tmp, index=False, engine=eng)
                used = eng
        else:
            eng = engine or _choose_engine()
            df.to_excel(tmp, index=False, engine=eng)
            used = eng
        os.replace(tmp, fn)
        logger.info("wrote file: %s (eng=%s) rows=%d", fn, used, len(df))
    except Exception:
        logger.exception("write file failed")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def _capture_debug_html(driver: FirefoxDriver, pref: str = "dbg") -> str:
    try:
        html = driver.page_source
        fd, tmp = tempfile.mkstemp(prefix=pref, suffix=".html")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(html)
        logger.debug("saved html: %s", tmp)
        return tmp
    except Exception:
        logger.exception("save debug html failed")
        return ""


def _ensure_dir(d: str) -> None:
    if not os.path.exists(d):
        os.makedirs(d, exist_ok=True)


def _switch_to_canvas_frame_if_any(driver: FirefoxDriver) -> bool:
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for idx, fr in enumerate(iframes):
            try:
                driver.switch_to.frame(fr)
                has = driver.execute_script(
                    "return !!document.querySelector('canvas');"
                )
                if has:
                    logger.debug("switched to iframe %d", idx)
                    return True
                driver.switch_to.default_content()
            except Exception:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                continue
    except Exception:
        logger.debug("iframe search error")
    return False


def _find_canvas_element(
    driver: FirefoxDriver, sels: List[str]
) -> Optional[WebElement]:
    for s in sels:
        try:
            el = driver.find_element(By.CSS_SELECTOR, s)
            if el:
                return el
        except NoSuchElementException:
            continue
        except Exception:
            continue
    try:
        return driver.find_element(By.CSS_SELECTOR, "canvas")
    except Exception:
        return None


def _capture_canvas_as_image(
    driver: FirefoxDriver, url: str, out_dir: str = "canvas_images"
) -> Optional[str]:
    try:
        _ensure_dir(out_dir)
        sels = [
            "canvas.jqplot-grid-canvas",
            "canvas.jqplot-event-canvas",
            "canvas.jqplot-canvas",
            "canvas.jqplot-overlay-canvas",
            "canvas",
        ]
        try:
            WebDriverWait(driver, EL_WAIT).until(
                lambda d: d.execute_script("return !!document.querySelector('canvas');")
            )
        except TimeoutException:
            logger.debug("canvas wait timeout %s", url)

        el = _find_canvas_element(driver, sels)
        if el is None:
            switched = _switch_to_canvas_frame_if_any(driver)
            if switched:
                el = _find_canvas_element(driver, sels)

        fname = os.path.join(out_dir, f"canvas_{abs(hash(url))}.png")

        if el is not None:
            try:
                png = el.screenshot_as_png
                if png:
                    with open(fname, "wb") as f:
                        f.write(png)
                    logger.debug("saved canvas png %s", fname)
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
                    return fname
            except Exception as e:
                logger.debug("screenshot failed: %s", e)

        try:
            canvas_data = driver.execute_script(
                "var selectors = ["
                " 'canvas.jqplot-grid-canvas',"
                " 'canvas.jqplot-event-canvas',"
                " 'canvas.jqplot-canvas',"
                " 'canvas.jqplot-overlay-canvas',"
                " 'canvas'"
                "];"
                "var c = null;"
                "for(var i=0;i<selectors.length;i++){"
                " c = document.querySelector(selectors[i]);"
                " if(c) break;"
                "}"
                "if(!c){"
                " function findCanvasInShadow(root){"
                "  if(!root) return null;"
                "  var nodes = root.querySelectorAll('*');"
                "  for(var i=0;i<nodes.length;i++){"
                "   var n = nodes[i];"
                "   if(n.shadowRoot){"
                "    var cc = n.shadowRoot.querySelector('canvas');"
                "    if(cc) return cc;"
                "    var d = findCanvasInShadow(n.shadowRoot);"
                "    if(d) return d;"
                "   }"
                "  }"
                "  return null;"
                " }"
                " c = findCanvasInShadow(document);"
                "}"
                "if(!c) return null;"
                "try { return c.toDataURL('image/png'); }"
                "catch(e) { return 'TAINTED:' + e.toString(); }"
            )
            if not canvas_data:
                logger.warning("no canvas js for %s", url)
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                return None
            if isinstance(canvas_data, str) and canvas_data.startswith("TAINTED:"):
                logger.warning("canvas tainted %s", url)
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                return None
            if isinstance(canvas_data, str) and canvas_data.startswith("data:"):
                b64 = canvas_data.split(",", 1)[1]
                img = base64.b64decode(b64)
                with open(fname, "wb") as f:
                    f.write(img)
                logger.debug("saved canvas via dataurl %s", fname)
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                return fname
        except Exception as e:
            logger.exception("toDataURL fallback err %s: %s", url, e)

        try:
            driver.switch_to.default_content()
        except Exception:
            pass

        logger.warning("capture failed %s", url)
        return None
    except Exception as e:
        logger.exception("capture canvas err %s: %s", url, e)
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return None


def _find_element_with_iframes(
    driver: FirefoxDriver, selector: str
) -> Optional[WebElement]:
    try:
        return driver.find_element(By.CSS_SELECTOR, selector)
    except Exception:
        pass
    try:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for idx, fr in enumerate(iframes):
            try:
                driver.switch_to.frame(fr)
                try:
                    el = driver.find_element(By.CSS_SELECTOR, selector)
                    if el:
                        logger.debug("found selector %s in iframe %d", selector, idx)
                        return el
                except Exception:
                    pass
                driver.switch_to.default_content()
            except Exception:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                continue
    except Exception:
        logger.debug("iframe search error for selector %s", selector)
    return None


def _capture_element_screenshot(
    driver: FirefoxDriver,
    selectors: List[str],
    url: str,
    out_dir: str = "canvas_images",
    prefix: str = "elem",
) -> Optional[str]:
    try:
        _ensure_dir(out_dir)
        fname = os.path.join(out_dir, f"{prefix}_{abs(hash(url))}.png")
        el = None
        for s in selectors:
            try:
                el = _find_element_with_iframes(driver, s)
                if el:
                    break
            except Exception:
                el = None
        if el is not None:
            try:
                png = el.screenshot_as_png
                if png:
                    with open(fname, "wb") as f:
                        f.write(png)
                    logger.debug("saved element png %s for selector %s", fname, s)
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
                    return fname
            except Exception as e:
                logger.debug("element screenshot failed: %s", e)

        try:
            full = os.path.join(out_dir, f"{prefix}_full_{abs(hash(url))}.png")
            driver.save_screenshot(full)
            logger.debug("saved full page fallback %s", full)
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            return full
        except Exception as e:
            logger.debug("full page screenshot failed: %s", e)
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            return None
    except Exception as e:
        logger.exception("capture element err %s: %s", url, e)
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return None


def extract_orders(
    driver: FirefoxDriver,
    url: str,
    retries: int = RETRY,
    out_dir: str = "canvas_images",
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Return paths: (graph_image, sell_table_image, buy_table_image)
    """
    try:
        driver.get(url)
    except WebDriverException:
        logger.warning("load err %s", url)
        return None, None, None
    time.sleep(1.0)
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    except Exception:
        pass
    time.sleep(1.0)

    attempt = 0
    back = 1.0
    while attempt <= retries:
        attempt += 1
        logger.debug("attempt %d/%d %s", attempt, retries + 1, url)
        try:
            WebDriverWait(driver, EL_WAIT).until(
                lambda d: d.execute_script(
                    "return !!(document.querySelector('#market_commodity_forsale_table') || "
                    "document.querySelector('#market_commodity_buyrequests_table') || "
                    "document.querySelector('#market_commodity_buyreqeusts_table') || "
                    "document.querySelector('canvas'));"
                )
            )
        except TimeoutException:
            logger.debug("wait for tables/canvas timeout %s", url)

        graph_path = _capture_canvas_as_image(driver, url, out_dir)

        sell_selectors = [
            "#market_commodity_forsale_table",
            "div#market_commodity_forsale_table",
            "div.market_commodity_orders_table_container#market_commodity_forsale_table",
            "div.market_commodity_orders_table_container table.market_commodity_orders_table",
        ]
        sell_path = _capture_element_screenshot(
            driver, sell_selectors, url, out_dir, prefix="sell"
        )

        buy_selectors = [
            "#market_commodity_buyrequests_table",
            "#market_commodity_buyreqeusts_table",
            "div#market_commodity_buyrequests_table",
            "div#market_commodity_buyreqeusts_table",
            "div.market_commodity_orders_table_container table.market_commodity_orders_table",
        ]
        buy_path = _capture_element_screenshot(
            driver, buy_selectors, url, out_dir, prefix="buy"
        )

        if graph_path or sell_path or buy_path:
            return graph_path, sell_path, buy_path

        if attempt <= retries:
            wt = back * RETRY_BACK
            logger.info("retry %d after %.1fs", attempt, wt)
            try:
                driver.refresh()
            except Exception:
                logger.debug("refresh failed")
            time.sleep(wt)
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            except Exception:
                pass
            time.sleep(1.0)
            back *= 2.0
            continue
        else:
            break

    dbg = _capture_debug_html(driver, pref="steam_dbg_")
    logger.warning("extract fail %s ; html %s", url, dbg)
    return None, None, None


def fixed_random_sleep(min_d: float = DELAY_MIN, max_d: float = DELAY_MAX) -> None:
    d = random.uniform(min_d, max_d)
    time.sleep(d)


def parse_args():
    p = argparse.ArgumentParser(description="Capture market screenshots.")
    p.add_argument("--input", "-i", default=IN_FILE)
    p.add_argument("--output", "-o", default=OUT_FILE)
    p.add_argument("--format", "-f", choices=["xlsx", "ods"], default="xlsx")
    p.add_argument("--engine", type=str, default=None)
    p.add_argument("--headless", action="store_true")
    p.add_argument("--test", action="store_true")
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--seed", type=int, default=None)
    p.add_argument("--no-open", action="store_true", default=True)
    p.add_argument("--verbose", action="store_true")
    p.add_argument("--user-agent", type=str, default=None)
    p.add_argument(
        "--no-random-ua",
        action="store_true",
        help="disable random UA",
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()
    logging.getLogger().setLevel(logging.DEBUG if args.verbose else logging.INFO)
    if args.seed is not None:
        random.seed(args.seed)

    ua = args.user_agent or (None if args.no_random_ua else random.choice(UAS))
    if ua:
        logger.debug("ua: %s", ua)

    input_rows = read_input_rows(args.input)
    if not input_rows:
        logger.info("no rows")
        return

    if args.test:
        input_rows = input_rows[:TEST_N]
    if args.limit and args.limit > 0:
        input_rows = input_rows[: args.limit]

    logger.info("processing %d rows", len(input_rows))

    out = args.output
    base, _ = os.path.splitext(out)
    out = base + (".ods" if args.format == "ods" else ".xlsx")
    prog = base + ".progress.json"

    visited, results = load_progress(prog)
    visited_urls = set(visited.keys())

    try:
        with firefox_driver(headless=args.headless, ua=ua) as drv:
            total = len(input_rows)
            for start in range(0, total, BATCH):
                end = min(start + BATCH, total)
                batch = input_rows[start:end]
                logger.info("batch %d -> %d", start + 1, end)
                for item in tqdm(batch, desc="visits", unit="pg"):
                    game, card, url = item
                    if not url or url.lower() in ("nan", "none"):
                        # keep a placeholder in visited to mark processed empty url
                        results.append((game, card, "", "", ""))
                        fixed_random_sleep()
                        continue
                    if url in visited_urls:
                        # already processed: do not re-run extraction
                        # ensure results keeps a record for compatibility
                        res = visited.get(url, None)
                        if res:
                            results.append(tuple(res))
                        else:
                            results.append((game, card, "", "", ""))
                        continue
                    try:
                        gpath, spath, bpath = extract_orders(drv, url)
                        row = (game, card, gpath or "", spath or "", bpath or "")
                        results.append(row)
                        # store authoritative row in visited
                        visited[url] = list(row)
                        visited_urls.add(url)
                        save_progress(prog, visited, results)
                    except Exception:
                        logger.exception("extract err %s", url)
                        results.append((game, card, "", "", ""))
                        visited[url] = [game, card, "", "", ""]
                        visited_urls.add(url)
                        save_progress(prog, visited, results)
                    fixed_random_sleep()
                # When saving intermediate output, build it from input_rows + visited
                save_output_from_input(input_rows, visited, out, engine=args.engine)
                time.sleep(2.0)
    except WebDriverException:
        logger.exception("webdriver global err")

    # Final save: authoritative output built from input_rows and visited
    save_output_from_input(input_rows, visited, out, engine=args.engine)
    save_progress(prog, visited, results)

    if visited:
        logger.info("done visited rows: %d", len(visited))
    else:
        logger.info("no results")


if __name__ == "__main__":
    main()
