#!/usr/bin/env python3
# steam_cards_sales_price_buyreqs_test50.py
"""
Lit steam_games_cards.xlsx et visite les Market URLs.
Sortie XLSX: A=Nom de la carte (col B input)
             B=Ventes
             C=Prix vendu
             D=Demandes
             E=Prix demandé

Traitement par lots de 100, délai aléatoire 10-15s.
Le fichier de sortie est ouvert une seule fois après
le premier lot si --no-open n'est pas fourni.
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
from typing import List, Tuple

import pandas as pd
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.webdriver import (
    WebDriver as FirefoxDriver,
)
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
from tqdm import tqdm

# Defaults
DEFAULT_INPUT = "steam_games_cards.xlsx"
DEFAULT_OUTPUT = "steam_games_cards_checked.xlsx"
URL_COLUMN = "Market URL"
CARD_COLUMN = "Nom de la carte"
PAGE_LOAD_TIMEOUT = 30

# Augmenter la valeur de test par défaut à 50 lignes
TEST_FIRST_N = 50

DELAY_MIN = 10.0
DELAY_MAX = 15.0
BATCH_SIZE = 100
ELEMENT_WAIT = 8

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("steam_cards_sales_price_buyreqs_test50")


def build_firefox_driver(headless: bool = False,
                         user_agent: str | None = None) -> WebDriver:
    opts = Options()
    if headless:
        opts.add_argument("--headless")
    if user_agent:
        opts.set_preference("general.useragent.override", user_agent)
    else:
        opts.set_preference(
            "general.useragent.override",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "Firefox/115.0",
        )
    gecko = GeckoDriverManager().install()
    svc = Service(executable_path=gecko)
    drv = FirefoxDriver(service=svc, options=opts)
    drv.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return drv


@contextmanager
def firefox_driver(headless: bool = False,
                   user_agent: str | None = None):
    drv = build_firefox_driver(headless=headless,
                               user_agent=user_agent)
    try:
        yield drv
    finally:
        try:
            drv.quit()
        except Exception:
            logger.exception("Erreur fermeture driver")


def read_input_rows(filename: str) -> List[Tuple[str, str]]:
    if not os.path.exists(filename):
        logger.error("Fichier introuvable: %s", filename)
        return []
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.exception("Impossible de lire %s", filename)
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
    rows: List[Tuple[str, str]] = []
    for i, u in enumerate(urls):
        card = cards[i] if i < len(cards) else ""
        rows.append((card, str(u).strip()))
    logger.info("Lignes lues: %d", len(rows))
    return rows


def save_output(rows: List[Tuple[str, str, str, str, str]],
                filename: str) -> None:
    df = pd.DataFrame(
        rows,
        columns=[
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


def extract_orders(driver: WebDriver, url: str
                   ) -> Tuple[str, str, str, str]:
    """
    Extrait:
      - ventes et prix depuis div#market_commodity_forsale
      - demandes et prix depuis div#market_commodity_buyrequests

    Retourne (sell_count, sell_price, buy_count, buy_price).
    Vide si absent.
    """
    try:
        driver.get(url)
    except WebDriverException:
        logger.warning("Erreur chargement %s", url)
        return "", "", "", ""
    wait = WebDriverWait(driver, ELEMENT_WAIT)

    sell_count = ""
    sell_price = ""
    try:
        sel = ("div#market_commodity_forsale"
               ".market_commodity_order_summary")
        el = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, sel)
        ))
    except Exception:
        try:
            el = driver.find_element(
                By.CSS_SELECTOR,
                "div#market_commodity_forsale"
                ".market_commodity_order_summary",
            )
        except Exception:
            el = None
    if el is not None:
        try:
            spans = el.find_elements(
                By.CSS_SELECTOR,
                "span.market_commodity_orders_header_promote",
            )
        except Exception:
            spans = []
        if len(spans) >= 1:
            try:
                sell_count = spans[0].text.strip()
            except Exception:
                sell_count = ""
        if len(spans) >= 2:
            try:
                sell_price = spans[1].text.strip()
            except Exception:
                sell_price = ""

    buy_count = ""
    buy_price = ""
    try:
        sel2 = ("div#market_commodity_buyrequests"
                ".market_commodity_order_summary")
        el2 = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, sel2)
        ))
    except Exception:
        try:
            el2 = driver.find_element(
                By.CSS_SELECTOR,
                "div#market_commodity_buyrequests"
                ".market_commodity_order_summary",
            )
        except Exception:
            el2 = None
    if el2 is not None:
        try:
            spans2 = el2.find_elements(
                By.CSS_SELECTOR,
                "span.market_commodity_orders_header_promote",
            )
        except Exception:
            spans2 = []
        if len(spans2) >= 1:
            try:
                buy_count = spans2[0].text.strip()
            except Exception:
                buy_count = ""
        if len(spans2) >= 2:
            try:
                buy_price = spans2[1].text.strip()
            except Exception:
                buy_price = ""

    return sell_count, sell_price, buy_count, buy_price


def fixed_random_sleep(min_d: float = DELAY_MIN,
                       max_d: float = DELAY_MAX) -> None:
    d = random.uniform(min_d, max_d)
    logger.info("Attente aléatoire: %.2f s", d)
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
    return p.parse_args()


def main() -> None:
    args = parse_args()
    logging.getLogger().setLevel(
        logging.DEBUG if args.verbose else logging.INFO
    )
    if args.seed is not None:
        random.seed(args.seed)

    rows_in = read_input_rows(args.input)
    if not rows_in:
        logger.info("Aucune ligne à traiter.")
        return

    if args.test:
        rows_in = rows_in[:TEST_FIRST_N]
    if args.limit and args.limit > 0:
        rows_in = rows_in[: args.limit]

    logger.info("Traitement: %d lignes", len(rows_in))

    results: List[Tuple[str, str, str, str, str]] = []
    first_open_done = False
    try:
        with firefox_driver(headless=args.headless,
                            user_agent=args.user_agent) as driver:
            total = len(rows_in)
            for start in range(0, total, BATCH_SIZE):
                end = min(start + BATCH_SIZE, total)
                batch = rows_in[start:end]
                logger.info("Traitement lot %d -> %d", start + 1, end)
                for item in tqdm(batch, desc="Visites", unit="page"):
                    card, url = item
                    if not url or url.lower() in ("nan", "none"):
                        results.append((card, "", "", "", ""))
                        fixed_random_sleep()
                        continue
                    try:
                        sell_c, sell_p, buy_c, buy_p = extract_orders(
                            driver, url
                        )
                        results.append((card, sell_c, sell_p, buy_c, buy_p))
                    except Exception:
                        logger.exception("Erreur extraction %s", url)
                        results.append((card, "", "", "", ""))
                    fixed_random_sleep()
                # Sauvegarde après chaque lot
                save_output(results, args.output)
                # Ouvrir le fichier une seule fois après premier lot
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
