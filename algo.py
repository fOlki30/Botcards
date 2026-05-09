import json
import re

driver.get("https://steamcommunity.com/market/listings/753/866510-Sylvia")

WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "pricehistory"))
)

soup = bs4.BeautifulSoup(driver.page_source, "html.parser")

price_history = []
for script in soup.find_all("script"):
    if script.string and "line1" in script.string:
        match = re.search(r"line1\s*=\s*(\[\[.+?\]\])", script.string, re.DOTALL)
        if match:
            data = json.loads(match.group(1))
            # Chaque entrée : [date, prix, nb_vendus]
            price_history = [
                {"date": entry[0], "price": entry[1], "sold": entry[2]}
                for entry in data
            ]
            break

print(price_history)

# -------------------------------------------------------------

driver.get("https://steamcommunity.com/market/listings/753/866510-Sylvia")

WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "pricehistory"))
)

soup = bs4.BeautifulSoup(driver.page_source, "html.parser")

table = soup.select("#market_commodity_forsale_table tr")[1:]  # ignore le header

offers = []
for row in table:
    cols = row.select("td")
    if len(cols) == 2:
        price = cols[0].get_text(strip=True)
        quantity = cols[1].get_text(strip=True)
        offers.append({"price": price, "quantity": quantity})

print(offers)