from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import pandas as pd
import time

driver = webdriver.Chrome(service=Service('D:/Selenium Project Work!/chromedriver.exe'), options=Options())

try:
    driver.get("https://www.psx.com.pk/market-summary/")

    time.sleep(10)  # wait 10 seconds for page and table to load

    rows = driver.find_elements(By.CSS_SELECTOR, "div.market-table__table__row")

    data = []
    for row in rows:
        cols = row.find_elements(By.CSS_SELECTOR, "div.market-table__table__cell")
        if len(cols) == 7:
            data.append([col.text.strip() for col in cols])

    df = pd.DataFrame(data, columns=["Symbol", "Open", "High", "Low", "Last", "Change", "Volume"])

    def to_float(x):
        try:
            return float(x.replace(",", ""))
        except:
            return None

    df["Last"] = df["Last"].apply(to_float)
    high_value_stocks = df[df["Last"] > 1000]

    high_value_stocks.to_excel("psx_high_value_stocks.xlsx", index=False)
    high_value_stocks.to_csv("psx_high_value_stocks.csv", index=False)

    print(f"Saved {len(high_value_stocks)} stocks priced above 1000 PKR.")

finally:
    driver.quit()
