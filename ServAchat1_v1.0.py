# acheter_cartes_playwright_firefox_v4.py
import time
import logging
import pandas as pd
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PTimeoutError,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

EXCEL = "Liste_cartes_a_acheter_test.xlsx"
SHEET = 0
NAME_IDX = 1  # colonne B -> index 1
PRICE_IDX = 3  # colonne D -> index 3
URL_IDX = 5  # colonne F -> index 5
HEADLESS = False
NAV_TO = 15000  # ms
SHORT = 0.5  # s

XPATH_ACH = "//span[contains(normalize-space(.),'Acheter')]"
SEL_PRICE = "#market_buy_commodity_input_price"
SEL_SSA = "#market_buyorder_dialog_accept_ssa"
XPATH_PLACE = "//span[contains(normalize-space(.)," "'PLACER L\\'ORDRE')]"
SEL_CLOSE = "div.newmodal_close"


def proc_row(page, name, url, price):
    logging.info("carte: %s", name)
    try:
        page.goto(url, timeout=NAV_TO)
    except Exception as e:
        logging.exception("nav fail %s: %s", url, e)
        return False

    try:
        page.wait_for_selector(XPATH_ACH, timeout=NAV_TO)
        page.click(XPATH_ACH)
        logging.info("clicked Acheter")
        time.sleep(SHORT)
    except PTimeoutError:
        logging.warning("Acheter not found %s", url)
        return False
    except Exception as e:
        logging.exception("click Acheter err %s: %s", url, e)
        return False

    try:
        page.wait_for_selector(SEL_PRICE, timeout=NAV_TO)
        page.fill(SEL_PRICE, str(price))
        logging.info("price filled %s", price)
        time.sleep(SHORT)
    except PTimeoutError:
        logging.warning("price input not found %s", url)
        return False
    except Exception as e:
        logging.exception("fill price err %s: %s", url, e)
        return False

    try:
        page.wait_for_selector(SEL_SSA, timeout=5000)
        try:
            checked = page.is_checked(SEL_SSA)
        except Exception:
            el = page.query_selector(SEL_SSA)
            checked = False
            if el:
                checked = page.eval_on_selector(SEL_SSA, "el => el.checked")
        if not checked:
            page.click(SEL_SSA)
            logging.info("SSA checked")
            time.sleep(SHORT)
        else:
            logging.info("SSA already checked")
    except PTimeoutError:
        logging.warning("SSA not found %s", url)
    except Exception as e:
        logging.exception("SSA err %s: %s", url, e)

    try:
        page.wait_for_selector(XPATH_PLACE, timeout=NAV_TO)
        page.click(XPATH_PLACE)
        logging.info("clicked PLACER L'ORDRE")
        time.sleep(1.0)
    except PTimeoutError:
        logging.warning("PLACER L'ORDRE not found %s", url)
        return False
    except Exception as e:
        logging.exception("place order err %s: %s", url, e)
        return False

    try:
        page.wait_for_selector(SEL_CLOSE, timeout=5000)
        page.click(SEL_CLOSE)
        logging.info("modal closed")
        time.sleep(SHORT)
    except PTimeoutError:
        logging.warning("modal close not found %s", url)
    except Exception as e:
        logging.exception("modal close err %s: %s", url, e)

    return True


def main():
    df = pd.read_excel(EXCEL, sheet_name=SHEET, engine="openpyxl")
    res = []
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=HEADLESS)
        ctx = browser.new_context()
        page = ctx.new_page()
        try:
            for i, row in df.iterrows():
                try:
                    name = row.iloc[NAME_IDX]
                    price = row.iloc[PRICE_IDX]
                    url = row.iloc[URL_IDX]
                except Exception:
                    logging.error("read fail line %s", i)
                    res.append(False)
                    continue

                if pd.isna(url):
                    logging.info("line %s empty url", i)
                    res.append(False)
                    continue

                ok = proc_row(page, name, str(url), price)
                res.append(ok)
                time.sleep(1.0)
        finally:
            ctx.close()
            browser.close()

    total = len(res)
    okc = sum(1 for r in res if r)
    logging.info("done %d/%d success", okc, total)


if __name__ == "__main__":
    main()
