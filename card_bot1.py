# card_bot_update_list_xlsx_tqdm.py
import errno
import logging
import os
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from urllib.parse import urljoin

from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.webdriver import (
    WebDriver as FirefoxDriver,
)
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager

import pandas as pd
from tqdm import tqdm

# Config
URL = "https://www.steamcardexchange.net/index.php?badgeprices"
BASE_URL = "https://www.steamcardexchange.net"
PAGE_LOAD_TIMEOUT = 20
ELEMENT_WAIT_TIMEOUT = 15
HEADLESS = False
OUTPUT_XLSX = "steam_games.xlsx"
TEMP_XLSX = OUTPUT_XLSX + ".tmp.xlsx"
DELAY_BETWEEN_ROWS = 0.02
COPY_RETRY_DELAY = 0.5
COPY_RETRY_COUNT = 6

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("card_bot_update_list_xlsx_tqdm")


def build_firefox_driver(headless: bool = HEADLESS) -> WebDriver:
    opts = Options()
    if headless:
        opts.add_argument("--headless")
    ua = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "Firefox/115.0"
    )
    opts.set_preference("general.useragent.override", ua)
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


def select_all_in_length_dropdown(driver: WebDriver) -> None:
    wait = WebDriverWait(driver, ELEMENT_WAIT_TIMEOUT)
    try:
        sel = wait.until(
            EC.presence_of_element_located(
                (By.NAME, "badgepricelist_guest_length")
            )
        )
        wait.until(
            EC.element_to_be_clickable(
                (By.NAME, "badgepricelist_guest_length")
            )
        )
        select = Select(sel)
        vals = [o.get_attribute("value") for o in select.options]
        if "-1" in vals:
            select.select_by_value("-1")
            logger.info("Option value='-1' sélectionnée (All).")
            time.sleep(1.0)
        else:
            logger.warning("Option value='-1' introuvable.")
    except TimeoutException:
        logger.exception("Timeout en attendant le select.")
        raise
    except NoSuchElementException:
        logger.exception("Élément introuvable.")
        raise


def extract_games_from_table(driver: WebDriver) -> list:
    """Retourne liste de tuples (nom, href)."""
    games = []
    wait = WebDriverWait(driver, ELEMENT_WAIT_TIMEOUT)
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody")))
        rows = driver.find_elements(
            By.CSS_SELECTOR, "tbody tr.odd, tbody tr.even"
        )
        logger.info("Lignes trouvées: %d", len(rows))
        for row in rows:
            try:
                name_elem = row.find_element(By.CSS_SELECTOR, "div.truncate")
                name = name_elem.text.strip()
            except NoSuchElementException:
                name = ""
            try:
                a_elem = row.find_element(
                    By.CSS_SELECTOR, "a[href*='gamepage']"
                )
                href = a_elem.get_attribute("href")
                if href and href.startswith("/"):
                    href = urljoin(BASE_URL, href)
            except NoSuchElementException:
                href = ""
            if name or href:
                games.append((name, href))
    except TimeoutException:
        logger.exception("Timeout en attendant le tableau.")
        raise
    return games


def read_existing_games_xlsx(filename: str) -> dict:
    """
    Lit un .xlsx existant et retourne dict key->(name, href).
    La clé est normalisée mais le nom conservé tel quel.
    """
    existing = {}
    if not os.path.exists(filename):
        return existing
    try:
        df = pd.read_excel(filename, engine="openpyxl")
        for _, row in df.iterrows():
            name = (
                str(row.iloc[0]).strip()
                if not pd.isna(row.iloc[0])
                else ""
            )
            href = (
                str(row.iloc[1]).strip()
                if len(row) > 1 and not pd.isna(row.iloc[1])
                else ""
            )
            key = normalize_key(name, href)
            if key:
                existing[key] = (name, href)
    except Exception:
        logger.warning("Impossible de lire %s; on repart de zéro.", filename)
    return existing


def save_games_xlsx(rows: list, filename: str, engine: str = "openpyxl"):
    """Écrit la liste de tuples (name, href) dans un .xlsx via pandas."""
    df = pd.DataFrame(rows, columns=["Nom du jeu", "URL"])
    dirn = os.path.dirname(os.path.abspath(filename)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dirn)
    os.close(fd)
    try:
        df.to_excel(tmp, index=False, engine=engine)
        os.replace(tmp, filename)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def normalize_key(name: str, href: str) -> str:
    """
    Retourne une clé stable pour fusion : nom normalisé en minuscule
    ou href si pas de nom.
    """
    if name:
        k = name.strip().lower()
        k = " ".join(k.split())
        return k
    return href.strip()


def merge_and_write_games(new_games: list) -> None:
    """
    Fusionne new_games avec le fichier existant sans doublons.
    Écrit d'abord dans TEMP_XLSX puis remplace OUTPUT_XLSX.
    Utilise tqdm pour afficher une barre de progression propre.
    """
    existing = read_existing_games_xlsx(OUTPUT_XLSX)
    for name, href in new_games:
        key = normalize_key(name, href)
        if not key:
            continue
        if href:
            # Conserver le nom tel qu'il apparaît dans la page HTML
            existing[key] = (name or existing.get(key, ("", ""))[0], href)
        else:
            existing.setdefault(key, existing.get(key, ("", "")))

    # existing: key -> (name, href)
    merged = sorted(existing.values(), key=lambda x: x[0].lower())
    rows = [(n, h) for n, h in merged]

    try:
        save_games_xlsx(rows, TEMP_XLSX)
        for attempt in range(COPY_RETRY_COUNT):
            try:
                os.replace(TEMP_XLSX, OUTPUT_XLSX)
                logger.info("Remplacement vers %s réussi.", OUTPUT_XLSX)
                break
            except OSError as exc:
                win32 = getattr(exc, "winerror", None) == 32
                eacces = getattr(exc, "errno", None) == errno.EACCES
                if win32 or eacces:
                    logger.warning(
                        "Tentative %d: %s est verrouillé, nouvelle tentative.",
                        attempt + 1,
                        OUTPUT_XLSX,
                    )
                    time.sleep(COPY_RETRY_DELAY)
                    continue
                logger.exception(
                    "Erreur inattendue lors du remplacement final."
                )
                break
        else:
            logger.warning(
                "Impossible de mettre à jour %s; vérifier manuellement.",
                OUTPUT_XLSX,
            )
    finally:
        total = len(rows)
        if total > 0:
            for _ in tqdm(
                range(total), desc="Écriture des lignes", unit="ligne"
            ):
                time.sleep(DELAY_BETWEEN_ROWS)
        try:
            open_file_with_default_app(OUTPUT_XLSX)
        except Exception:
            logger.info("Impossible d'ouvrir %s automatiquement.", OUTPUT_XLSX)


def open_file_with_default_app(path: str) -> None:
    """Ouvre le fichier avec l'application par défaut (non bloquant)."""
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


def main() -> None:
    try:
        logger.info("Lancement du navigateur Firefox.")
        with firefox_driver(headless=HEADLESS) as driver:
            driver.get(URL)
            logger.info("Page chargée : %s", driver.title)
            select_all_in_length_dropdown(driver)
            logger.info("Etape 1 terminée: 'All' sélectionné.")
            games = extract_games_from_table(driver)
            logger.info("Jeux extraits: %d", len(games))
            print("Mise à jour de la liste des jeux...")
            merge_and_write_games(games)
            print("Terminé. Le fichier final a été ouvert si possible.")
    except WebDriverException as exc:
        logger.exception("Erreur WebDriver détectée.", exc_info=exc)
    finally:
        logger.info("Script terminé.")


if __name__ == "__main__":
    main()
