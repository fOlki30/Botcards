#!/usr/bin/env python3
# steam_history_bot_pages.py

from __future__ import annotations

import argparse
import logging
import os
import random
import sys
import time
from contextlib import contextmanager
from typing import List, Dict, Optional

import pandas as pd
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.webdriver import WebDriver as FirefoxDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

logger = logging.getLogger("steam_history_bot")

STEAM_MARKET_URL = "https://steamcommunity.com/market/"
DEFAULT_OUTPUT = "steam_market_history.xlsx"
PAGE_TIMEOUT = 60
EL_WAIT = 20
MAX_PAGES = 5


def build_driver(
    headless: bool = False,
    ua: Optional[str] = None
) -> FirefoxDriver:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.set_preference("dom.webdriver.enabled", False)
    opts.set_preference("useAutomationExtension", False)
    if ua:
        opts.set_preference(
            "general.useragent.override",
            ua
        )
    gecko = GeckoDriverManager().install()
    svc = Service(executable_path=gecko)
    drv = FirefoxDriver(service=svc, options=opts)
    drv.set_page_load_timeout(PAGE_TIMEOUT)
    return drv


@contextmanager
def firefox_driver(
    headless: bool = False,
    ua: Optional[str] = None,
):
    drv = build_driver(headless=headless, ua=ua)
    try:
        yield drv
    finally:
        try:
            drv.quit()
        except Exception:
            logger.exception("error closing driver")


def ask_continue() -> bool:
    """Demande à l'utilisateur de continuer [Y/N]."""
    while True:
        try:
            ans = input(
                "Autoriser le bot à cliquer sur l'onglet historique ? [Y/N]: "
            ).strip()
        except EOFError:
            return False
        if not ans:
            continue
        a = ans[0].lower()
        if a == "y":
            return True
        if a == "n":
            return False
        print("Répondez par Y ou N.")


def click_market_history_tab(
    driver: FirefoxDriver
) -> bool:
    """Clique l'onglet 'Mon historique sur le marché'."""
    xpaths = [
        "//span[normalize-space() = 'Mon "
        "historique sur le marché']",
        "//span[contains(normalize-space(), 'Mon "
        "historique')]",
        "//a[contains(@class,'market_tab') "
        "and contains(., 'Mon historique')]",
    ]
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, EL_WAIT).until(
                EC.element_to_be_clickable(
                    (By.XPATH, xp)
                )
            )
            driver.execute_script(
                "arguments[0].scrollIntoView(true);", el
            )
            time.sleep(0.3)
            el.click()
            logger.info("clicked history tab via xpath")
            return True
        except (TimeoutException, WebDriverException,
                NoSuchElementException):
            continue
        except Exception:
            continue

    try:
        els = driver.find_elements(
            By.CSS_SELECTOR,
            "span.market_tab_well_tab_contents"
        )
        for el in els:
            try:
                txt = el.text or ""
                if "historique" in txt.lower():
                    driver.execute_script(
                        "arguments[0].scrollIntoView(true);",
                        el
                    )
                    time.sleep(0.3)
                    el.click()
                    logger.info("clicked history tab via CSS")
                    return True
            except Exception:
                continue
    except Exception:
        pass

    logger.warning("could not find or click the history tab")
    return False


def wait_for_history_rows(
    driver: FirefoxDriver,
    timeout: int = EL_WAIT
) -> bool:
    """Attendre la présence probable de lignes d'historique."""
    selectors = [
        "div.market_listing_row",
        "div.market_listing_row_link",
        "div.market_listing_table",
    ]
    for sel in selectors:
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.find_elements(
                    By.CSS_SELECTOR,
                    sel
                )
            )
            els = driver.find_elements(
                By.CSS_SELECTOR,
                sel
            )
            if els:
                logger.debug(
                    "found rows for selector: %s",
                    sel
                )
                return True
        except TimeoutException:
            continue
        except Exception:
            continue
    return False


def parse_history_rows(
    driver: FirefoxDriver
) -> List[Dict[str, str]]:
    """
    Parcourt les lignes visibles et extrait:
    Nom du jeu, nom de la carte, gain, prix, dates.
    Filtre les cartes contenant "carte à échanger" dans
    le nom du jeu.
    """
    rows: List[Dict[str, str]] = []

    candidates = driver.find_elements(
        By.CSS_SELECTOR,
        "div.market_listing_row, div.market_listing_row_link"
    )
    if not candidates:
        candidates = driver.find_elements(
            By.CSS_SELECTOR,
            "div.market_listing_table div"
        )

    for it in candidates:
        try:
            # gain or loss
            try:
                g_el = it.find_element(
                    By.CSS_SELECTOR,
                    "div.market_listing_left_cell."
                    "market_listing_gainorloss"
                )
                gain = g_el.text.strip()
            except Exception:
                gain = ""

            # price
            try:
                p_el = it.find_element(
                    By.CSS_SELECTOR,
                    "div.market_listing_right_cell."
                    "market_listing_their_price "
                    ".market_listing_price"
                )
                price = p_el.text.strip()
            except Exception:
                try:
                    p_el = it.find_element(
                        By.CSS_SELECTOR,
                        ".market_listing_price"
                    )
                    price = p_el.text.strip()
                except Exception:
                    price = ""

            # dates
            date1 = ""
            date2 = ""
            try:
                date_els = it.find_elements(
                    By.CSS_SELECTOR,
                    "div.market_listing_right_cell."
                    "market_listing_listed_date.can_combine"
                )
                if date_els:
                    if len(date_els) >= 1:
                        date1 = date_els[0].text.strip()
                    if len(date_els) >= 2:
                        date2 = date_els[1].text.strip()
            except Exception:
                pass

            if not date2:
                try:
                    comb = it.find_element(
                        By.CSS_SELECTOR,
                        "div.market_listing_listed_date_combined"
                    )
                    txt = comb.text.strip()
                    if txt:
                        parts = txt.split(":")
                        if len(parts) >= 2:
                            date2 = parts[1].strip()
                        else:
                            date2 = txt
                except Exception:
                    pass

            # item name
            card = ""
            try:
                name_el = it.find_element(
                    By.CSS_SELECTOR,
                    "span.market_listing_item_name"
                )
                card = name_el.text.strip()
            except Exception:
                try:
                    name_spans = it.find_elements(
                        By.CSS_SELECTOR,
                        "span"
                    )
                    for sp in name_spans:
                        sid = sp.get_attribute("id") or ""
                        sid_prefix = "history_row_"
                        sid_suffix = "_name"
                        ok1 = sid.startswith(sid_prefix)
                        ok2 = sid.endswith(sid_suffix)
                        if ok1 and ok2:
                            card = sp.text.strip()
                            break
                except Exception:
                    pass

            # game name
            game = ""
            try:
                game_el = it.find_element(
                    By.CSS_SELECTOR,
                    "span.market_listing_game_name"
                )
                game = game_el.text.strip()
            except Exception:
                try:
                    gdiv = it.find_element(
                        By.CSS_SELECTOR,
                        "div.market_listing_item_name_block"
                    )
                    spans = gdiv.find_elements(
                        By.CSS_SELECTOR,
                        "span"
                    )
                    if spans:
                        for sp in spans:
                            txt = sp.text.strip()
                            if txt and "carte" in txt.lower():
                                game = txt
                                break
                except Exception:
                    pass

            # Filtre les cartes
            if "carte à échanger" not in game.lower():
                continue

            card = card or ""
            game = game or ""
            gain = gain or ""
            price = price or ""
            date1 = date1 or ""
            date2 = date2 or ""

            if card or price:
                rows.append({
                    "Nom du jeu": game,
                    "Nom de la carte": card,
                    "GainOrLoss": gain,
                    "Prix": price,
                    "Conclusion": date1,
                    "Mise en ligne": date2,
                })
        except Exception:
            continue

    logger.info("parsed %d history rows", len(rows))
    return rows


def click_next_page(driver: FirefoxDriver) -> bool:
    """
    Clique le bouton '>' pour aller à la page suivante.
    Retourne True si le clic a été effectué.
    """
    try:
        btn = driver.find_element(
            By.ID,
            "tabContentsMyMarketHistory_btn_next"
        )
    except Exception:
        return False

    try:
        dis = btn.get_attribute("disabled")
        if dis and dis.lower() in ("true", "disabled"):
            return False
        cls = btn.get_attribute("class") or ""
        if "disabled" in cls.lower():
            return False
        aria = btn.get_attribute("aria-disabled")
        if aria and aria.lower() == "true":
            return False

        time.sleep(random.uniform(3.0, 5.0))
        driver.execute_script(
            "arguments[0].scrollIntoView(true);",
            btn
        )
        time.sleep(0.2)
        btn.click()
        logger.info("clicked next page button")
        return True
    except Exception:
        logger.debug("click next failed")
        return False


def save_to_excel(
    rows: List[Dict[str, str]],
    out_path: str,
    append: bool = False
) -> None:
    if not rows:
        logger.info("no rows to save")
        return
    df = pd.DataFrame(rows)
    cols = [
        "Nom du jeu",
        "Nom de la carte",
        "GainOrLoss",
        "Prix",
        "Conclusion",
        "Mise en ligne",
    ]
    df = df.reindex(columns=cols)
    try:
        if append and os.path.exists(out_path):
            existing_df = pd.read_excel(out_path)
            df = pd.concat(
                [existing_df, df],
                ignore_index=True
            )
        df.to_excel(out_path, index=False)
        logger.info("wrote output file: %s", out_path)
    except Exception as e:
        logger.exception(
            "failed to write output file: %s",
            e
        )


def main() -> None:
    p = argparse.ArgumentParser(
        description="Steam market history extractor"
    )
    p.add_argument(
        "--headless",
        action="store_true",
        help="run browser headless"
    )
    p.add_argument(
        "--output",
        "-o",
        default=DEFAULT_OUTPUT,
        help="output xlsx file"
    )
    p.add_argument(
        "--user-agent",
        default=None,
        help="override user agent string"
    )
    p.add_argument(
        "--append",
        action="store_true",
        help="append to existing file if it exists"
    )
    p.add_argument(
        "--max-pages",
        type=int,
        default=MAX_PAGES,
        help=f"maximum number of pages to scrape (default: {MAX_PAGES})"
    )
    args = p.parse_args()

    all_rows: List[Dict[str, str]] = []

    try:
        logger.info("starting bot...")
        with firefox_driver(
            headless=args.headless,
            ua=args.user_agent
        ) as drv:
            try:
                logger.info("opening %s", STEAM_MARKET_URL)
                drv.get(STEAM_MARKET_URL)

                # Attendre que la page soit chargée
                time.sleep(5.0)
                logger.info("page loaded")

                cont = ask_continue()
                if not cont:
                    logger.info("user chose not to continue; exiting")
                    sys.exit(0)

                logger.info("proceeding to click history tab")
                clicked = click_market_history_tab(drv)
                if not clicked:
                    logger.warning("history tab not clicked")

                wait_for_history_rows(drv, timeout=EL_WAIT)
                time.sleep(1.0)

                page = 1
                while page <= args.max_pages:
                    logger.info("parsing page %d", page)
                    rows = parse_history_rows(drv)
                    if rows:
                        all_rows.extend(rows)

                    if page >= args.max_pages:
                        logger.info("max pages reached: %d", args.max_pages)
                        break
                    moved = click_next_page(drv)
                    if not moved:
                        logger.info("no next page or end reached")
                        break

                    time.sleep(1.0)
                    wait_for_history_rows(drv, timeout=EL_WAIT)
                    time.sleep(0.8)
                    page += 1

                save_to_excel(
                    all_rows,
                    args.output,
                    append=args.append
                )
            except WebDriverException as e:
                logger.exception("failed to load page: %s", e)
            except Exception as e:
                logger.exception("unexpected error: %s", e)

    except Exception as e:
        logger.exception("bot encountered an error: %s", e)

    logger.info("done")


if __name__ == "__main__":
    main()
