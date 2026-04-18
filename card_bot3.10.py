#!/usr/bin/env python3
# steam_cards_sales_price_buyreqs_robust_wrapped_fixed.py
"""
Lit steam_games_cards.xlsx et visite les Market URLs.
Sortie XLSX: A=Nom du jeu (col A input) B=Nom de la carte (col B input)
C=Ventes D=Prix vendu E=Demandes F=Prix demandé
Version robuste, lignes <= 79 caractères.
"""
from __future__ import annotations

import argparse
import logging
import os
import random
import sys
import tempfile
import time
from contextlib import contextmanager
from typing import List, Tuple, Optional

import pandas as pd
from selenium.common.exceptions import (
    WebDriverException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.webdriver import (
    WebDriver as FirefoxDriver,
)
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
from tqdm import tqdm

DEFAULT_INPUT = "steam_games_cards.xlsx"
DEFAULT_OUTPUT = "steam_games_cards_prices.xlsx"
URL_COLUMN = "Market URL"
CARD_COLUMN = "Nom de la carte"
GAME_COLUMN = "Nom du jeu"
PAGE_LOAD_TIMEOUT = 60

TEST_FIRST_N = 20
DELAY_MIN = 7.0
DELAY_MAX = 10.0
BATCH_SIZE = 100

ELEMENT_WAIT = 20
RETRY_COUNT = 3
RETRY_BACKOFF = 2.0

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("steam_cards_robust_wrapped_fixed")

DEFAULT_USER_AGENTS = [
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


def build_firefox_driver(headless: bool = False,
                         user_agent: Optional[str] = None) -> WebDriver:
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.set_preference("dom.webdriver.enabled", False)
    opts.set_preference("useAutomationExtension", False)
    if user_agent:
        opts.set_preference("general.useragent.override", user_agent)
    gecko = GeckoDriverManager().install()
    svc = Service(executable_path=gecko)
    drv = FirefoxDriver(service=svc, options=opts)
    drv.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return drv


@contextmanager
def firefox_driver(headless: bool = False,
                   user_agent: Optional[str] = None):
    drv = build_firefox_driver(headless=headless, user_agent=user_agent)
    try:
        yield drv
    finally:
        try:
            drv.quit()
        except Exception:
            logger.exception("Erreur fermeture driver")


def read_input_rows(filename: str) -> List[Tuple[str, str, str]]:
    if not os.path.exists(filename):
        logger.error("Fichier introuvable: %s", filename)
        return []
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.exception("Impossible de lire %s", filename)
        return []
    if GAME_COLUMN in df.columns:
        games = df[GAME_COLUMN].astype(str).tolist()
    elif df.shape[1] >= 1:
        games = df.iloc[:, 0].astype(str).tolist()
    else:
        logger.error("Colonne nom du jeu introuvable.")
        return []
    if CARD_COLUMN in df.columns:
        cards = df[CARD_COLUMN].astype(str).tolist()
    elif df.shape[1] >= 2:
        cards = df.iloc[:, 1].astype(str).tolist()
    else:
        logger.error("Colonne carte introuvable.")
        return []
    if URL_COLUMN in df.columns:
        urls = df[URL_COLUMN].astype(str).tolist()
    elif df.shape[1] >= 3:
        urls = df.iloc[:, 2].astype(str).tolist()
    else:
        logger.error("Colonne URL introuvable.")
        return []
    rows: List[Tuple[str, str, str]] = []
    for i, u in enumerate(urls):
        game = games[i] if i < len(games) else ""
        card = cards[i] if i < len(cards) else ""
        rows.append((str(game).strip(), str(card).strip(), str(u).strip()))
    logger.info("Lignes lues: %d", len(rows))
    return rows


def save_output(rows: List[Tuple[str, str, str, str, str, str]],
                filename: str) -> None:
    df = pd.DataFrame(
        rows,
        columns=[
            "Nom du jeu",
            "Nom de la carte",
            "Ventes",
            "Prix vendu",
            "Demandes",
            "Prix demandé",
        ],
    )
    dirn = os.path.dirname(os.path.abspath(filename)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dirn)
    os.close(fd)
    try:
        df.to_excel(tmp, index=False, engine="openpyxl")
        os.replace(tmp, filename)
        logger.info("Fichier écrit: %s", filename)
    except Exception:
        logger.exception("Erreur écriture fichier")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def open_file(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            import subprocess
            subprocess.Popen(["open", path])
        else:
            import subprocess
            try:
                subprocess.Popen(["xdg-open", path])
            except FileNotFoundError:
                logger.info("xdg-open non dispo; fichier sauvé.")
    except Exception:
        logger.exception("Impossible d'ouvrir %s", path)


def _capture_debug_html(driver: WebDriver, prefix: str = "debug") -> str:
    try:
        html = driver.page_source
        fd, tmp = tempfile.mkstemp(prefix=prefix, suffix=".html")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(html)
        logger.debug("HTML debug sauvegardé: %s", tmp)
        return tmp
    except Exception:
        logger.exception("Impossible de sauvegarder HTML debug")
        return ""


def _wait_for_order_summary(driver: WebDriver,
                            timeout: float = ELEMENT_WAIT) -> Optional[Tuple]:
    wait = WebDriverWait(driver, timeout)
    sel_sell = ("div#market_commodity_forsale."
                "market_commodity_order_summary")
    sel_buy = ("div#market_commodity_buyrequests."
               "market_commodity_order_summary")
    try:
        wait.until(
            lambda d: d.execute_script(
                "return !!(document.querySelector(arguments[0]) || "
                "document.querySelector(arguments[1]));",
                sel_sell, sel_buy
            )
        )
    except TimeoutException:
        return None
    try:
        el_sell = driver.find_element(By.CSS_SELECTOR, sel_sell)
    except Exception:
        el_sell = None
    try:
        el_buy = driver.find_element(By.CSS_SELECTOR, sel_buy)
    except Exception:
        el_buy = None
    return el_sell, el_buy


def _extract_from_summary_element(el) -> Tuple[str, str]:
    if el is None:
        return "", ""
    try:
        spans = el.find_elements(
            By.CSS_SELECTOR,
            "span.market_commodity_orders_header_promote",
        )
    except Exception:
        spans = []
    count = spans[0].text.strip() if len(spans) >= 1 else ""
    price = spans[1].text.strip() if len(spans) >= 2 else ""
    return count, price


def extract_orders(driver: WebDriver, url: str,
                   retries: int = RETRY_COUNT) -> Tuple[str, str, str, str]:
    try:
        driver.get(url)
    except WebDriverException:
        logger.warning("Erreur chargement initial %s", url)
        return "", "", "", ""
    time.sleep(1.0)
    try:
        script = (
            "window.scrollTo(0, document.body.scrollHeight);"
        )
        driver.execute_script(script)
    except Exception:
        pass
    time.sleep(1.0)

    attempt = 0
    backoff = 1.0
    while attempt <= retries:
        attempt += 1
        logger.debug("Extraction attempt %d/%d pour %s",
                     attempt, retries + 1, url)
        found = _wait_for_order_summary(driver, timeout=ELEMENT_WAIT)
        if found is not None:
            el_sell, el_buy = found
            sell_count, sell_price = _extract_from_summary_element(el_sell)
            buy_count, buy_price = _extract_from_summary_element(el_buy)
            if sell_count or buy_count or sell_price or buy_price:
                logger.debug("Extraction réussie au attempt %d", attempt)
                return sell_count, sell_price, buy_count, buy_price
            else:
                logger.debug(
                    "Order summary présent mais sans spans utiles "
                    "(attempt %d)", attempt
                )
        else:
            logger.debug("Order summary non présent (attempt %d)", attempt)

        if attempt <= retries:
            wait_time = backoff * RETRY_BACKOFF
            logger.info("Retry %d pour %s après %.1fs (refresh)",
                        attempt, url, wait_time)
            try:
                driver.refresh()
            except Exception:
                logger.debug("Refresh échoué, tentative suivante")
            time.sleep(wait_time)
            try:
                script = (
                    "window.scrollTo(0, document.body.scrollHeight);"
                )
                driver.execute_script(script)
            except Exception:
                pass
            time.sleep(1.0)
            backoff *= 2.0
            continue
        else:
            break

    debug_path = _capture_debug_html(driver, prefix="steam_debug_")
    logger.warning("Échec extraction pour %s ; HTML sauvegardé: %s",
                   url, debug_path)
    return "", "", "", ""


def fixed_random_sleep(min_d: float = DELAY_MIN,
                       max_d: float = DELAY_MAX) -> None:
    d = random.uniform(min_d, max_d)
    time.sleep(d)


def parse_args():
    p = argparse.ArgumentParser(
        description="Extrait ventes, prix, demandes et prix demandés."
    )
    p.add_argument("--input", "-i", default=DEFAULT_INPUT)
    p.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    p.add_argument("--headless", action="store_true")
    p.add_argument("--test", action="store_true")
    p.add_argument("--limit", type=int, default=0)
    p.add_argument("--seed", type=int, default=None)
    p.add_argument("--no-open", action="store_true")
    p.add_argument("--verbose", action="store_true")
    p.add_argument("--user-agent", type=str, default=None)
    p.add_argument("--no-random-ua", action="store_true",
                   help="Désactive la sélection d'un UA aléatoire")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    logging.getLogger().setLevel(
        logging.DEBUG if args.verbose else logging.INFO
    )
    if args.seed is not None:
        random.seed(args.seed)

    ua = None
    if args.user_agent:
        ua = args.user_agent
    elif not args.no_random_ua:
        ua = random.choice(DEFAULT_USER_AGENTS)
        logger.debug("User-Agent aléatoire utilisé: %s", ua)

    rows_in = read_input_rows(args.input)
    if not rows_in:
        logger.info("Aucune ligne à traiter.")
        return

    if args.test:
        rows_in = rows_in[:TEST_FIRST_N]
    if args.limit and args.limit > 0:
        rows_in = rows_in[: args.limit]

    logger.info("Traitement: %d lignes", len(rows_in))

    results: List[Tuple[str, str, str, str, str, str]] = []
    first_open_done = False
    try:
        with firefox_driver(headless=args.headless, user_agent=ua) as driver:
            total = len(rows_in)
            for start in range(0, total, BATCH_SIZE):
                end = min(start + BATCH_SIZE, total)
                batch = rows_in[start:end]
                logger.info("Traitement lot %d -> %d", start + 1, end)
                for item in tqdm(batch, desc="Visites", unit="page"):
                    game, card, url = item
                    if not url or url.lower() in ("nan", "none"):
                        results.append((game, card, "", "", "", ""))
                        fixed_random_sleep()
                        continue
                    try:
                        sell_c, sell_p, buy_c, buy_p = extract_orders(
                            driver, url
                        )
                        results.append(
                            (game, card, sell_c, sell_p, buy_c, buy_p)
                        )
                    except Exception:
                        logger.exception("Erreur extraction %s", url)
                        results.append((game, card, "", "", "", ""))
                    fixed_random_sleep()
                save_output(results, args.output)
                if not first_open_done and not args.no_open:
                    try:
                        open_file(args.output)
                        first_open_done = True
                    except Exception:
                        logger.info("Impossible d'ouvrir le fichier.")
                time.sleep(2.0)
    except WebDriverException:
        logger.exception("Erreur WebDriver globale.")

    if results:
        logger.info("Terminé, total lignes: %d", len(results))
    else:
        logger.info("Aucun résultat extrait.")


if __name__ == "__main__":
    main()
