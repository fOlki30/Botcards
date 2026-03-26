#!/usr/bin/env python3
# card_bot_test_toggle.py
"""
Extractor: par défaut traite toutes les URLs. Si --test est fourni,
le script limite le traitement aux 10 premières URLs.
Sortie XLSX: A Nom du jeu | B Nom de la carte | C Market URL
Usage: python card_bot_test_toggle.py --input steam_games.xlsx
"""
import argparse
import errno
import logging
import os
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from typing import List, Tuple
import re

import pandas as pd
from tqdm import tqdm
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

# Configuration
DEFAULT_INPUT = "steam_games.xlsx"
DEFAULT_OUTPUT = "steam_games_cards.xlsx"
URL_COLUMN = "URL"
PAGE_LOAD_TIMEOUT = 30
ELEMENT_WAIT_TIMEOUT = 12
HEADLESS = False
DELAY_BETWEEN_PAGES = 0.6
COPY_RETRY_DELAY = 0.5
COPY_RETRY_COUNT = 6
TEST_FIRST_N = 10  # nombre de pages en mode test

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("card_bot_test_toggle")


def build_firefox_driver(headless: bool = HEADLESS) -> WebDriver:
    opts = Options()
    if headless:
        opts.add_argument("--headless")
    opts.set_preference(
        "general.useragent.override",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Firefox/115.0",
    )
    gecko = GeckoDriverManager().install()
    svc = Service(executable_path=gecko)
    drv = FirefoxDriver(service=svc, options=opts)
    drv.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return drv


@contextmanager
def firefox_driver(headless: bool = HEADLESS):
    drv = build_firefox_driver(headless=headless)
    try:
        yield drv
    finally:
        try:
            drv.quit()
        except Exception:
            logger.exception("Erreur lors de la fermeture du driver")


def read_urls_from_xlsx(filename: str, col: str = URL_COLUMN) -> List[str]:
    if not os.path.exists(filename):
        logger.error("Fichier introuvable: %s", filename)
        return []
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.exception("Impossible de lire %s", filename)
        return []
    if col in df.columns:
        urls = df[col].astype(str).tolist()
    elif df.shape[1] >= 1:
        urls = df.iloc[:, 0].astype(str).tolist()
    else:
        logger.error("Colonne %s introuvable dans %s", col, filename)
        return []
    cleaned = [u.strip() for u in urls if isinstance(u, str) and u.strip()]
    logger.info("URLs lues: %d", len(cleaned))
    return cleaned


def _find_market_in_container(cont) -> str:
    try:
        pri = cont.find_elements(
            By.CSS_SELECTOR, "a.mt-auto.btn-primary[href]"
        )
    except Exception:
        pri = []
    for a in pri:
        try:
            href = a.get_attribute("href") or ""
            if "market/listings" in href:
                return href
        except Exception:
            continue
    try:
        all_a = cont.find_elements(By.CSS_SELECTOR, "a[href]")
    except Exception:
        all_a = []
    for a in all_a:
        try:
            href = a.get_attribute("href") or ""
            if "market/listings" in href:
                return href
        except Exception:
            continue
    return ""


# Exclusion patterns pour éléments non-cartes
_EXCLUDE_PATTERNS = [
    re.compile(r"Profile\s*Background", re.I),
    re.compile(r"%3A.*%3A"),
    re.compile(r"Chat\s*Preview", re.I),
    re.compile(r"Profile", re.I),
    re.compile(r"Sticker", re.I),
    re.compile(r"Background", re.I),
]


def _market_url_is_card(href: str) -> bool:
    if not href:
        return False
    href_l = href.lower()
    # Exclure explicitement les backgrounds
    if "profile%20background" in href_l or "profile background" in href_l:
        return False
    for pat in _EXCLUDE_PATTERNS:
        if pat.search(href):
            return False
    # mentions explicites
    if "trading%20card" in href_l:
        return True
    if "trading card" in href_l:
        return True
    if "(foil)" in href_l or "%20(foil)" in href_l:
        return True
    m = re.search(r"/market/listings/[^/]+/([^/?#]+)", href)
    if not m:
        return False
    tail = m.group(1)
    if "%3A" in tail:
        return False
    if re.search(r"\(.*(Profile|Background).*?\)", tail, re.I):
        return False
    if re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", tail):
        return True
    return False


def _is_card_container(cont) -> bool:
    try:
        img = cont.find_element(By.CSS_SELECTOR, "img[data-gallery-type]")
        dtype = (img.get_attribute("data-gallery-type") or "").lower()
        if "card" in dtype:
            return True
    except Exception:
        pass
    try:
        img2 = cont.find_element(By.CSS_SELECTOR, "img[data-gallery-desc]")
        desc = (img2.get_attribute("data-gallery-desc") or "").lower()
        if "card" in desc or "series" in desc:
            return True
    except Exception:
        pass
    try:
        href = _find_market_in_container(cont)
        if href and _market_url_is_card(href):
            return True
    except Exception:
        pass
    return False


# Regex to remove leading "Series ... - Card ... - " prefixes
_PREFIX_RE = re.compile(
    r"^(?:Series\s*\d+\s*-\s*)?"
    r"(?:Card\s*\d+(?:\s*of\s*\d+)?\s*-\s*)+",
    flags=re.I,
)


def clean_card_name(raw: str) -> str:
    if not raw:
        return ""
    name = raw.strip()
    name = re.sub(r"\s*-\s*", " - ", name)
    prev = None
    while prev != name:
        prev = name
        name = _PREFIX_RE.sub("", name).strip()
    name = name.replace("(Foil)", "").strip()
    return name


def extract_from_page(driver: WebDriver, url: str
                      ) -> List[Tuple[str, str, str]]:
    rows: List[Tuple[str, str, str]] = []
    try:
        driver.get(url)
    except WebDriverException:
        logger.exception("Erreur lors du chargement de %s", url)
        return rows

    wait = WebDriverWait(driver, ELEMENT_WAIT_TIMEOUT)

    game_name = ""
    selectors = [
        "span.font-semibold.truncate",
        "h1",
        "h2",
        "div.game_title",
    ]
    for sel in selectors:
        try:
            el = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, sel))
            )
            game_name = el.text.strip()
            if game_name:
                break
        except Exception:
            try:
                el = driver.find_element(By.CSS_SELECTOR, sel)
                game_name = el.text.strip()
                if game_name:
                    break
            except Exception:
                continue

    card_container_sel = (
        "div.flex.flex-col.items-center.p-5.gap-y-2.bg-gray-light"
    )
    try:
        conts = driver.find_elements(By.CSS_SELECTOR, card_container_sel)
    except Exception:
        conts = []

    if not conts:
        try:
            conts = driver.find_elements(
                By.CSS_SELECTOR,
                "div[data-gallery-desc], div.gallery-image-anchor",
            )
        except Exception:
            conts = []

    for cont in conts:
        try:
            if not _is_card_container(cont):
                continue

            card_name = ""
            try:
                img = cont.find_element(
                    By.CSS_SELECTOR, "img[data-gallery-desc]"
                )
                card_name = img.get_attribute("data-gallery-desc") or ""
            except Exception:
                pass

            if not card_name:
                try:
                    img2 = cont.find_element(By.CSS_SELECTOR, "img[alt]")
                    card_name = img2.get_attribute("alt") or ""
                except Exception:
                    pass

            if not card_name:
                try:
                    name_elem = cont.find_element(
                        By.CSS_SELECTOR,
                        "div.text-sm.text-center.break-words",
                    )
                    card_name = name_elem.text.strip()
                except Exception:
                    raw = cont.text.strip()
                    card_name = raw.splitlines()[0] if raw else ""

            # Nettoyage des préfixes "Series ... - Card ... - "
            card_name = clean_card_name(card_name)

            # Exclure si le nom indique background/profile
            cn_l = card_name.lower()
            if "background" in cn_l or "profile" in cn_l:
                continue

            market_url = _find_market_in_container(cont)
            if not market_url or not _market_url_is_card(market_url):
                continue

            # Marquer Premium si foil dans l'URL
            if "(Foil)" in market_url or "foil" in market_url.lower():
                if not card_name.endswith(" (Premium)"):
                    card_name = f"{card_name} (Premium)"

            rows.append((game_name, card_name, market_url))
        except Exception:
            logger.exception("Erreur lors de l'extraction d'un conteneur")
            continue

    if not rows:
        rows.append((game_name, "", ""))

    return rows


def save_rows_to_xlsx(rows: List[Tuple[str, str, str]],
                      filename: str) -> None:
    df = pd.DataFrame(
        rows,
        columns=[
            "Nom du jeu",
            "Nom de la carte",
            "Market URL",
        ],
    )
    dirn = os.path.dirname(os.path.abspath(filename)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dirn)
    os.close(fd)
    try:
        df.to_excel(tmp, index=False, engine="openpyxl")
        for attempt in range(COPY_RETRY_COUNT):
            try:
                os.replace(tmp, filename)
                logger.info("Fichier écrit: %s", filename)
                break
            except OSError as exc:
                win32 = getattr(exc, "winerror", None) == 32
                eacces = getattr(exc, "errno", None) == errno.EACCES
                if win32 or eacces:
                    logger.warning(
                        "Tentative %d: %s verrouillé, nouvelle tentative.",
                        attempt + 1,
                        filename,
                    )
                    time.sleep(COPY_RETRY_DELAY)
                    continue
                logger.exception("Erreur inattendue lors de l'écriture.")
                break
        else:
            logger.warning("Impossible d'écrire %s; vérifier manuellement.",
                           filename)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def open_file_with_default_app(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", path])
        else:
            try:
                subprocess.Popen(["xdg-open", path])
            except FileNotFoundError:
                subprocess.Popen(["soffice", "--calc", path])
    except Exception:
        logger.exception("Impossible d'ouvrir le fichier %s", path)


def parse_args():
    p = argparse.ArgumentParser(
        description="Traitement standard : --test active la limitation."
    )
    p.add_argument("--input", "-i", default=DEFAULT_INPUT)
    p.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    p.add_argument("--headless", action="store_true")
    p.add_argument("--test", action="store_true",
                   help="Activer le mode test (limite à 10 pages).")
    p.add_argument("--limit", type=int, default=0,
                   help="Limiter le nombre de pages (0 = toutes).")
    p.add_argument("--no-open", action="store_true",
                   help="Ne pas ouvrir le fichier de sortie.")
    p.add_argument("--verbose", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    urls = read_urls_from_xlsx(args.input, URL_COLUMN)
    if not urls:
        logger.info("Aucune URL à traiter, sortie.")
        return

    # Par défaut : traiter toutes les URLs.
    # Si --test est fourni, limiter aux TEST_FIRST_N premiers.
    if args.test:
        urls = urls[:TEST_FIRST_N]
    if args.limit and args.limit > 0:
        urls = urls[: args.limit]

    logger.info("Traitement (mode test=%s): %d pages",
                args.test, len(urls))

    all_rows: List[Tuple[str, str, str]] = []
    try:
        with firefox_driver(headless=args.headless) as driver:
            for url in tqdm(urls, desc="Pages", unit="page"):
                try:
                    rows = extract_from_page(driver, url)
                    all_rows.extend(rows)
                    time.sleep(DELAY_BETWEEN_PAGES)
                except Exception:
                    logger.exception("Erreur lors du traitement de %s", url)
                    continue
    except WebDriverException:
        logger.exception("Erreur WebDriver globale.")

    if all_rows:
        save_rows_to_xlsx(all_rows, args.output)
        if not args.no_open:
            try:
                open_file_with_default_app(args.output)
            except Exception:
                logger.info("Impossible d'ouvrir %s automatiquement.",
                            args.output)
    else:
        logger.info("Aucun résultat extrait.")


if __name__ == "__main__":
    main()
