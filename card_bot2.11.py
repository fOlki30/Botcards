# -*- coding: utf-8 -*-

"""
card_bot_improved_fixed_premium.py
Version modifiée pour :
 - éviter d'enregistrer des nombres/compteurs comme noms de cartes
 - extraire un nom depuis l'URL market si le texte visible est suspect
 - détecter correctement les cartes Premium/Foil depuis l'URL et les attributs
 - sauvegarder snippets HTML pour debug
 - heuristiques renforcées pour trouver le nom
   (data attrs, alt, title, aria-label)
Usage: python card_bot_improved_fixed_premium.py --input steam_games.xlsx
"""

import argparse
import errno
import logging
import os
import subprocess
import sys
import tempfile
import time
import json
import re
from contextlib import contextmanager
from typing import List, Tuple, Optional
from urllib.parse import unquote

import pandas as pd
from tqdm import tqdm
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.webdriver import WebDriver as FirefoxDriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager

# Configuration par défaut
DEFAULT_INPUT = "steam_games.xlsx"
DEFAULT_OUTPUT = "steam_games_cards.xlsx"
PAGE_LOAD_TIMEOUT = 45
ELEMENT_WAIT_TIMEOUT = 18
HEADLESS = False
DELAY_BETWEEN_PAGES = 0.8
COPY_RETRY_DELAY = 0.6
COPY_RETRY_COUNT = 6
TEST_FIRST_N = 10

# Heuristiques
SCROLL_PAUSE = 0.8
MAX_SCROLL_ITER = 40
LOAD_MORE_SELECTORS = [
    "button.load-more",
    "button[data-action='load-more']",
    "a.load-more",
    "button#load_more",
]

# Logging
logger = logging.getLogger("card_bot_improved_fixed_premium")

# -------------------------
# Driver helpers
# -------------------------


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


# -------------------------
# I/O Excel
# -------------------------

def read_input_xlsx(filename: str) -> List[Tuple[str, str, int]]:
    if not os.path.exists(filename):
        logger.error("Fichier input introuvable : %s", filename)
        return []
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.exception("Impossible de lire %s", filename)
        return []

    rows: List[Tuple[str, str, int]] = []
    for _, row in df.iterrows():
        game = str(row.get("Nom du jeu", "") or "").strip()
        url = str(row.get("URL", "") or "").strip()
        raw_status = row.get("Statut", 1)
        try:
            status = int(raw_status)
        except Exception:
            status = 1
        rows.append((game, url, status))

    logger.info("Lignes d'input lues: %d", len(rows))
    return rows


def read_existing_output_xlsx(filename: str) -> List[Tuple[str, str, str]]:
    rows: List[Tuple[str, str, str]] = []
    if not os.path.exists(filename):
        return rows
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.warning(
            "Impossible de lire %s ; les données existantes seront ignorées.",
            filename,
        )
        return rows
    for _, row in df.iterrows():
        game = (
            str(row.iloc[0]).strip()
            if len(row) > 0 and not pd.isna(row.iloc[0])
            else ""
        )
        card = (
            str(row.iloc[1]).strip()
            if len(row) > 1 and not pd.isna(row.iloc[1])
            else ""
        )
        market = (
            str(row.iloc[2]).strip()
            if len(row) > 2 and not pd.isna(row.iloc[2])
            else ""
        )
        if game or card or market:
            rows.append((game, card, market))
    return rows


def save_rows_to_xlsx(rows: List[Tuple[str, str, str]], filename: str) -> None:
    df = pd.DataFrame(
        rows,
        columns=["Nom du jeu", "Nom de la carte", "Market URL"],
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
            logger.warning(
                "Impossible d'écrire %s; vérifier manuellement.",
                filename,
            )
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


# -------------------------
# Utilitaires d'extraction
# -------------------------

_EXCLUDE_PATTERNS = [
    re.compile(r"Profile\s*Background", re.I),
    re.compile(r"%3A.*%3A"),
    re.compile(r"Chat\s*Preview", re.I),
    re.compile(r"Profile", re.I),
    re.compile(r"Sticker", re.I),
    re.compile(r"Background", re.I),
]

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
    name = re.sub(r"\(Foil\)|\(foil\)", "", name, flags=re.I).strip()
    return name


def decode_href(href: str) -> str:
    try:
        return unquote(href)
    except Exception:
        return href or ""


def _market_url_is_card(href: str) -> bool:
    if not href:
        return False
    href_dec = decode_href(href).lower()
    # Exclure explicitement les backgrounds
    if "profile background" in href_dec or "profile%20background" in href_dec:
        return False
    for pat in _EXCLUDE_PATTERNS:
        if pat.search(href):
            return False
    # mentions explicites
    if "trading%20card" in href_dec or "trading card" in href_dec:
        return True
    if "(foil)" in href_dec or " foil" in href_dec:
        return True
    # extraire la partie finale du listing
    m = re.search(r"/market/listings/[^/]+/([^/?#]+)", href_dec)
    if not m:
        # si pas de pattern, mais contient 'market/listings' et des lettres,
        # on accepte
        if "market/listings" in href_dec and re.search(r"[a-zà-ÿ]", href_dec):
            return True
        return False
    tail = m.group(1)
    if "%3A" in tail:
        return False
    if re.search(r"\(.*(Profile|Background).*?\)", tail, re.I):
        return False
    if re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", tail):
        return True
    # si la queue contient uniquement des chiffres ou chiffres-tiret-chiffres,
    # on l'accepte aussi (ex: 591420-4)
    if re.fullmatch(r"[\d\-_%]+", tail):
        return True
    return False


def _is_card_container(cont) -> bool:
    # heuristiques multiples : attributs data, alt text, liens market
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
        # fallback : alt text
        img3 = cont.find_element(By.CSS_SELECTOR, "img[alt]")
        alt = (img3.get_attribute("alt") or "").lower()
        if "card" in alt or "foil" in alt:
            return True
    except Exception:
        pass
    try:
        href = _find_market_in_container(cont)
        if href and _market_url_is_card(href):
            return True
    except Exception:
        pass
    # dernier recours : texte contenant "card" ou "foil"
    try:
        txt = cont.text or ""
        if re.search(r"\b(card|foil|trading)\b", txt, re.I):
            return True
    except Exception:
        pass
    return False


def _find_market_in_container(cont) -> str:
    try:
        pri = cont.find_elements(
            By.CSS_SELECTOR,
            "a.mt-auto.btn-primary[href], a[href].market_link",
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


def try_extract_market_endpoints_from_page_source(source: str) -> List[str]:
    """
    Cherche dans le HTML/JS des endpoints JSON ou des objets contenant
    'market/listings' ou des structures JSON utiles.
    """
    endpoints = set()
    # pattern simple pour URLs market
    for m in re.finditer(
        r"https?://store\.steampowered\.com/market/listings/[^\"'\\s]+",
        source,
    ):
        endpoints.add(m.group(0))
    # pattern pour endpoints JSON (ex: /market/priceoverview/)
    for m in re.finditer(r"(/market/priceoverview/\?[^\"'\\s]+)", source):
        endpoints.add("https://store.steampowered.com" + m.group(1))
    return list(endpoints)


# -------------------------
# Validation et extraction depuis URL
# -------------------------

def looks_like_valid_name(s: str) -> bool:
    if not s:
        return False
    s = s.strip()
    if len(s) < 2:
        return False
    # doit contenir au moins une lettre
    if re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", s):
        return True
    # accepter aussi un petit nombre (ex: "4") si c'est raisonnable
    if re.fullmatch(r"\d{1,3}", s):
        return True
    return False


def extract_name_from_market_url(href: str) -> str:
    """
    Extrait un candidat de nom depuis la dernière partie d'une URL Steam
    Market.
    Exemples:
      .../market/listings/753/591420-4  -> "4"
      .../market/listings/753/591420-My-Card-Name -> "My Card Name"
    """
    if not href:
        return ""
    try:
        href_dec = decode_href(href)
        tail = href_dec.rstrip("/").split("/")[-1]
        tail = re.split(r"[?#]", tail)[0]
        tail = tail.replace("%3A", ":")
        tail = tail.replace("_", " ").replace("%20", " ").strip()
        if "-" in tail:
            parts = tail.split("-")
            if re.fullmatch(r"\d+", parts[0]) and len(parts) >= 2:
                candidate = "-".join(parts[1:])
            else:
                candidate = "-".join(parts)
        else:
            candidate = tail
        candidate = re.sub(r"^\-+|\-+$", "", candidate).strip()
        candidate = re.sub(r"%\w{2}", " ", candidate)
        candidate = candidate.replace("-", " ").strip()
        candidate = re.sub(r"\s+", " ", candidate)
        return candidate
    except Exception:
        return ""


# -------------------------
# Premium / Foil detection
# -------------------------

def is_premium_from_market_url(href: str) -> bool:
    """
    Retourne True si l'URL de marché contient un indicateur 'foil'/'premium'.
    Gère les formes encodées (%28%29, %20), les parenthèses,
    les query params, etc.
    """
    if not href:
        return False
    try:
        dec = decode_href(href).lower()
    except Exception:
        dec = (href or "").lower()

    # recherche simple mot 'foil' ou 'premium'
    if re.search(r"\bfoil\b", dec):
        return True
    if re.search(r"\bpremium\b", dec):
        return True

    # formes entre parenthèses : (foil), (Foil) encodées ou non
    if re.search(r"\(.*foil.*\)", dec):
        return True
    if "%28" in href.lower() and "%29" in href.lower() and "foil" in dec:
        return True

    # segments de path contenant '-foil' ou 'foil-' ou '_foil'
    if re.search(r"[-_/]foil[-_/]?", dec):
        return True

    # query params ou fragments contenant foil
    if re.search(r"[?&#][^#]*foil", dec):
        return True

    return False


def is_premium_from_container(cont) -> bool:
    """
    Inspecte data-gallery-desc, alt, title, aria-label et texte visible
    pour détecter 'foil' ou 'premium'.
    """
    try:
        # data-gallery-desc
        try:
            el = cont.find_element(By.CSS_SELECTOR, "img[data-gallery-desc]")
            desc = (el.get_attribute("data-gallery-desc") or "").lower()
            if re.search(r"\bfoil\b|\bpremium\b", desc):
                return True
        except Exception:
            pass

        # alt/title/aria-label
        for attr in ("alt", "title", "aria-label"):
            try:
                el = cont.find_element(By.CSS_SELECTOR, f"[{attr}]")
                val = (el.get_attribute(attr) or "").lower()
                if re.search(r"\bfoil\b|\bpremium\b", val):
                    return True
            except Exception:
                pass

        # texte visible
        try:
            txt = (cont.text or "").lower()
            if re.search(r"\bfoil\b|\bpremium\b", txt):
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


# -------------------------
# Lazy load helpers
# -------------------------

def scroll_to_load_all(driver: WebDriver,
                       container_selector: Optional[str] = None,
                       debug_dir: Optional[str] = None) -> None:
    """
    Scroll progressif pour forcer le lazy-loading.
    Si container_selector fourni,
    on observe le nombre d'enfants pour savoir quand s'arrêter.
    """
    same_count_iter = 0
    for i in range(MAX_SCROLL_ITER):
        try:
            if container_selector:
                elems = driver.find_elements(By.CSS_SELECTOR,
                                             container_selector)
                cur_count = len(elems)
            else:
                elems = driver.find_elements(By.CSS_SELECTOR, "img")
                cur_count = len(elems)
        except Exception:
            cur_count = -1

        try:
            driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight);")
        except Exception:
            pass

        time.sleep(SCROLL_PAUSE)

        for sel in LOAD_MORE_SELECTORS:
            try:
                btns = driver.find_elements(By.CSS_SELECTOR, sel)
                for b in btns:
                    try:
                        if b.is_displayed() and b.is_enabled():
                            logger.debug("Clicking load more button: %s", sel)
                            driver.execute_script("arguments[0].click();", b)
                            time.sleep(SCROLL_PAUSE)
                    except Exception:
                        continue
            except Exception:
                continue

        try:
            if container_selector:
                elems2 = driver.find_elements(By.CSS_SELECTOR,
                                              container_selector)
                new_count = len(elems2)
            else:
                elems2 = driver.find_elements(By.CSS_SELECTOR, "img")
                new_count = len(elems2)
        except Exception:
            new_count = -1

        if new_count == cur_count:
            same_count_iter += 1
        else:
            same_count_iter = 0

        if same_count_iter >= 3:
            logger.debug("No new elements after scrolling (iter=%d).",
                         same_count_iter)
            break

    if debug_dir:
        try:
            os.makedirs(debug_dir, exist_ok=True)
            idx = int(time.time())
            path = os.path.join(debug_dir,
                                f"page_after_scroll_{idx}.html")
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(driver.page_source)
            logger.debug("Saved page source to %s", path)
        except Exception:
            logger.exception("Impossible de sauvegarder le HTML de debug.")


# -------------------------
# Extraction principale
# -------------------------

def extract_from_page(driver: WebDriver, url: str,
                      debug_dir: Optional[str] = None
                      ) -> List[Tuple[str, str, str]]:
    rows: List[Tuple[str, str, str]] = []
    try:
        driver.get(url)
    except WebDriverException:
        logger.exception("Erreur lors du chargement de %s", url)
        return rows

    wait = WebDriverWait(driver, ELEMENT_WAIT_TIMEOUT)

    # attendre un titre raisonnable
    game_name = ""
    selectors = ["span.font-semibold.truncate", "h1", "h2",
                 "div.game_title", "div.apphub_AppName"]
    for sel in selectors:
        try:
            el = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, sel)))
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

    # Forcer le lazy-load : scroll + click load more
    card_container_sel_candidates = [
        "div.flex.flex-col.items-center.p-5.gap-y-2.bg-gray-light",
        "div[data-gallery-desc]",
        "div.gallery-image-anchor",
        "div.card_item",
        "div.card",
        "div.text-sm.text-center.break-words"
    ]
    found_container_sel = None
    for sel in card_container_sel_candidates:
        try:
            if driver.find_elements(By.CSS_SELECTOR, sel):
                found_container_sel = sel
                break
        except Exception:
            continue

    scroll_to_load_all(driver,
                       container_selector=found_container_sel,
                       debug_dir=debug_dir)

    # tenter d'extraire des endpoints market depuis le source
    endpoints = []
    try:
        src = driver.page_source or ""
        endpoints = try_extract_market_endpoints_from_page_source(src)
        if debug_dir and endpoints:
            with open(os.path.join(debug_dir, "endpoints_found.json"),
                      "w", encoding="utf-8") as fh:
                json.dump(endpoints, fh, ensure_ascii=False, indent=2)
    except Exception:
        logger.debug(
            "Erreur lors de l'extraction d'endpoints depuis le source.")

    # collecter conteneurs potentiels
    conts = []
    try:
        for sel in card_container_sel_candidates:
            try:
                found = driver.find_elements(By.CSS_SELECTOR, sel)
                if found:
                    conts.extend(found)
            except Exception:
                continue
        if not conts:
            conts = driver.find_elements(
                By.CSS_SELECTOR,
                ("img[data-gallery-desc], div.gallery-image-anchor, "
                 "div.card_item, div.card"))
    except Exception:
        conts = []

    rejected_hrefs = []
    for cont in conts:
        try:
            if not _is_card_container(cont):
                try:
                    a = cont.find_element(By.CSS_SELECTOR, "a[href]")
                    href = a.get_attribute("href") or ""
                    if href:
                        rejected_hrefs.append(href)
                except Exception:
                    pass
                continue

            card_name = ""
            # 1) tenter data attributes et alt/title/aria-label
            try:
                img = cont.find_element(By.CSS_SELECTOR,
                                        "img[data-gallery-desc]")
                card_name = img.get_attribute("data-gallery-desc") or ""
            except Exception:
                pass

            if not card_name:
                try:
                    img2 = cont.find_element(By.CSS_SELECTOR, "img[alt]")
                    card_name = img2.get_attribute("alt") or ""
                except Exception:
                    pass

            # 2) heuristique renforcée pour récupérer le nom
            if not card_name:
                candidates = []
                try:
                    candidates = cont.find_elements(
                        By.CSS_SELECTOR,
                        (
                            "img[data-gallery-desc], img[alt],"
                            " [data-gallery-desc], [title], [aria-label],"
                            " .card_name, .card-title,"
                            " div.text-sm.text-center.break-words"
                        )
                    )
                except Exception:
                    candidates = []

                chosen = ""
                for el in candidates:
                    try:
                        txt = (
                            el.get_attribute("data-gallery-desc")
                            or el.get_attribute("alt")
                            or el.get_attribute("title")
                            or el.get_attribute("aria-label")
                            or el.text
                            or ""
                        ).strip()
                        if looks_like_valid_name(txt):
                            chosen = txt
                            break
                    except Exception:
                        continue

                # fallback: texte visible non numérique le plus long
                if not chosen:
                    try:
                        texts = [
                            t.strip()
                            for t in cont.text.splitlines()
                            if t.strip()
                        ]
                        texts = [
                            t
                            for t in texts
                            if looks_like_valid_name(t)
                        ]
                        if texts:
                            chosen = max(texts, key=len)
                    except Exception:
                        chosen = ""

                # si toujours rien ou suspect,
                # tenter d'extraire depuis le market URL
                if not looks_like_valid_name(chosen):
                    try:
                        market_url = _find_market_in_container(cont)
                        if market_url:
                            candidate_from_url = (
                                extract_name_from_market_url(market_url)
                            )
                            if looks_like_valid_name(candidate_from_url):
                                chosen = candidate_from_url
                    except Exception:
                        pass

                # dernier recours: inspecter images
                if not chosen:
                    try:
                        img = cont.find_element(
                            By.CSS_SELECTOR,
                            "img[data-gallery-desc], img[alt]"
                        )
                        chosen = (
                            img.get_attribute("data-gallery-desc")
                            or img.get_attribute("alt")
                            or ""
                        ).strip()
                    except Exception:
                        pass

                card_name = chosen

            # récupérer market_url tôt (avant nettoyage)
            market_url = _find_market_in_container(cont)
            if not market_url and endpoints:
                for ep in endpoints:
                    if "market/listings" in ep:
                        market_url = ep
                        break

            # detect premium BEFORE cleaning name
            is_premium = False
            try:
                if market_url and is_premium_from_market_url(market_url):
                    is_premium = True
                elif is_premium_from_container(cont):
                    is_premium = True
            except Exception:
                logger.debug(
                    "Erreur lors de la détection premium pour le "
                    "conteneur."
                )

            # clean name
            card_name = clean_card_name(card_name)
            cn_l = card_name.lower()
            if "background" in cn_l or "profile" in cn_l:
                continue

            if not market_url or not _market_url_is_card(market_url):
                try:
                    a = cont.find_element(By.CSS_SELECTOR, "a[href]")
                    href = a.get_attribute("href") or ""
                    if href:
                        rejected_hrefs.append(href)
                except Exception:
                    pass
                continue

            # annotate premium
            if is_premium and card_name:
                if not card_name.endswith(" (Premium)"):
                    card_name = f"{card_name} (Premium)"

            # debug : si on a extrait un nom numérique depuis l'URL,
            # sauvegarder snippet
            try:
                if debug_dir and re.fullmatch(r"\d{1,3}", card_name):
                    os.makedirs(debug_dir, exist_ok=True)
                    ts = int(time.time())
                    snippet = (
                        cont.get_attribute("outerHTML")
                        or cont.get_attribute("innerHTML")
                        or cont.text
                    )
                    with open(
                        os.path.join(debug_dir, f"card_from_url_{ts}.html"),
                        "w",
                        encoding="utf-8",
                    ) as fh:
                        fh.write(f"market_url: {market_url}\n\n")
                        fh.write(snippet)
                    logger.debug(
                        "Saved container snippet used to extract numeric "
                        "card name: card_from_url_%d.html",
                        ts,
                    )
            except Exception:
                logger.exception(
                    "Impossible de sauvegarder le conteneur debug."
                )

            rows.append((game_name, card_name, market_url))
        except Exception:
            logger.exception("Erreur lors de l'extraction d'un conteneur")
            continue

    if not rows:
        rows.append((game_name, "", ""))

    if debug_dir:
        try:
            os.makedirs(debug_dir, exist_ok=True)
            ts = int(time.time())
            rej_path = os.path.join(debug_dir, f"rejected_hrefs_{ts}.txt")
            with open(rej_path, "w", encoding="utf-8") as fh:
                for h in rejected_hrefs:
                    fh.write(h + "\n")
            logger.debug("Saved rejected hrefs to %s", rej_path)
        except Exception:
            logger.exception("Impossible d'écrire les hrefs rejetés.")

    return rows


# -------------------------
# Consistency check (premium vs normal)
# -------------------------

def check_premium_cards_consistency(
    rows: List[Tuple[str, str, str]]
) -> Tuple[List[str], List[Tuple[str, str, str]]]:
    cards_by_game = {}
    for game, card, url in rows:
        if game not in cards_by_game:
            cards_by_game[game] = {"normal": 0, "premium": 0, "rows": []}
        cards_by_game[game]["rows"].append((game, card, url))
        if "(Premium)" in card:
            cards_by_game[game]["premium"] += 1
        elif card:
            cards_by_game[game]["normal"] += 1

    discordant_games = []
    missing_rows = []
    for game, counts in cards_by_game.items():
        normal_count = counts["normal"]
        premium_count = counts["premium"]
        if normal_count != premium_count:
            logger.warning("%s: %d cartes normales vs %d premium.",
                           game, normal_count, premium_count)
            discordant_games.append(game)
            missing_rows.extend(counts["rows"])
    if discordant_games:
        logger.info("Jeux avec asymétrie carte/premium: %s",
                    ", ".join(discordant_games))
    return discordant_games, missing_rows


# -------------------------
# Input file updates
# -------------------------

def remove_status_2_from_input(filename: str) -> None:
    if not os.path.exists(filename):
        return
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.warning("Impossible de lire %s pour suppression.", filename)
        return
    if "Statut" in df.columns:
        df_filtered = df[df["Statut"] != 2]
        removed_count = len(df) - len(df_filtered)
    else:
        logger.warning("Colonne 'Statut' introuvable dans %s.", filename)
        return
    if removed_count == 0:
        logger.info("Aucune ligne avec statut 2 à supprimer.")
        return
    dirn = os.path.dirname(os.path.abspath(filename)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dirn)
    os.close(fd)
    try:
        df_filtered.to_excel(tmp, index=False, engine="openpyxl")
        os.replace(tmp, filename)
        logger.info("Fichier input mis à jour : %d lignes supprimées.",
                    removed_count)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def update_status_1_to_0(filename: str) -> None:
    if not os.path.exists(filename):
        return
    try:
        df = pd.read_excel(filename, engine="openpyxl")
    except Exception:
        logger.warning("Impossible de lire %s pour mise à jour.", filename)
        return
    if "Statut" not in df.columns:
        logger.warning("Colonne 'Statut' introuvable dans %s.", filename)
        return
    updated_count = (df["Statut"] == 1).sum()
    if updated_count == 0:
        logger.info("Aucun statut 1 à mettre à jour.")
        return
    df.loc[df["Statut"] == 1, "Statut"] = 0
    dirn = os.path.dirname(os.path.abspath(filename)) or "."
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dirn)
    os.close(fd)
    try:
        df.to_excel(tmp, index=False, engine="openpyxl")
        os.replace(tmp, filename)
        logger.info("Fichier input mis à jour : %d statuts 1 changés en 0.",
                    updated_count)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


# -------------------------
# Main
# -------------------------

def parse_args():
    p = argparse.ArgumentParser(
        description="Traitement amélioré : lazy-load, heuristiques, debug.")
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
    p.add_argument("--debug", action="store_true",
                   help="Sauvegarde HTML et hrefs rejetés pour debug.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s")

    input_rows = read_input_xlsx(args.input)
    if not input_rows:
        logger.info("Aucune ligne d'input à traiter, sortie.")
        return

    game_order = {}
    for idx, (game, url, status) in enumerate(input_rows):
        if game not in game_order:
            game_order[game] = idx

    existing_output = read_existing_output_xlsx(args.output)
    existing_by_game = {}
    for row in existing_output:
        existing_by_game.setdefault(row[0], []).append(row)

    to_process: List[Tuple[str, str]] = []
    final_rows: List[Tuple[str, str, str]] = []
    removed_games = set()

    for game, url, status in input_rows:
        if status == 2:
            if game:
                removed_games.add(game)
            continue
        if status == 0:
            if game in existing_by_game:
                final_rows.extend(existing_by_game[game])
            else:
                if url:
                    logger.warning(
                        ("Statut 0 pour %s sans données existantes. "
                         "Récupération des données manquantes."),
                        game)
                    to_process.append((game, url))
                else:
                    logger.warning(
                        "Statut 0 pour %s sans URL et sans données.",
                        game)
            continue
        if status == 1:
            if not url:
                logger.warning(
                    "Statut 1 pour %s sans URL, impossible de mettre à jour.",
                    game)
                continue
            to_process.append((game, url))
            continue
        logger.warning("Statut inconnu %s pour %s, traitement ignoré.",
                       status, game)

    if args.test:
        to_process = to_process[:TEST_FIRST_N]
    if args.limit and args.limit > 0:
        to_process = to_process[: args.limit]

    logger.info("Traitement (mode test=%s): %d pages",
                args.test, len(to_process))

    debug_dir = None
    if args.debug:
        debug_dir = os.path.join(os.getcwd(), "card_bot_debug")
        os.makedirs(debug_dir, exist_ok=True)

    try:
        with firefox_driver(headless=args.headless) as driver:
            for game, url in tqdm(to_process, desc="Pages", unit="page"):
                try:
                    rows = extract_from_page(driver, url, debug_dir=debug_dir)
                    final_rows.extend(rows)
                    time.sleep(DELAY_BETWEEN_PAGES)
                except Exception:
                    logger.exception("Erreur lors du traitement de %s", url)
                    continue
    except WebDriverException:
        logger.exception("Erreur WebDriver globale.")

    # Vérifier la cohérence des cartes premium vs normales
    discordant_games, missing_rows = check_premium_cards_consistency(
        final_rows)

    if discordant_games:
        logger.info("Rescrapage des jeux discordants: %s",
                    ", ".join(discordant_games))
        games_to_rescrape: List[Tuple[str, str]] = []
        for game, url, status in input_rows:
            if game in discordant_games and url:
                games_to_rescrape.append((game, url))

        final_rows = [row for row in final_rows
                      if row[0] not in discordant_games]

        try:
            with firefox_driver(headless=args.headless) as driver:
                for game, url in tqdm(games_to_rescrape,
                                      desc="Rescrapage", unit="page"):
                    try:
                        rows = extract_from_page(driver, url,
                                                 debug_dir=debug_dir)
                        final_rows.extend(rows)
                        time.sleep(DELAY_BETWEEN_PAGES)
                    except Exception:
                        logger.exception("Erreur lors du rescrapage de %s",
                                         url)
                        continue
        except WebDriverException:
            logger.exception("Erreur WebDriver lors du rescrapage.")

    remove_status_2_from_input(args.input)
    update_status_1_to_0(args.input)

    final_rows.sort(key=lambda row: game_order.get(row[0], float("inf")))

    if final_rows:
        save_rows_to_xlsx(final_rows, args.output)
        if not args.no_open:
            try:
                open_file_with_default_app(args.output)
            except Exception:
                logger.exception("Impossible d'ouvrir le fichier de sortie.")


if __name__ == "__main__":
    main()
