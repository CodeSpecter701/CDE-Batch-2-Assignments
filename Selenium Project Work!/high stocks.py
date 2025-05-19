from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

driver_path = 'D:/Selenium Project Work!/chromedriver.exe'
service = Service(driver_path)
options = Options()

driver = webdriver.Chrome(service=service, options=options)

try:
    driver.get("https://www.psx.com.pk/market-summary/")

    WebDriverWait(driver, 40).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.market-table__table__row"))
    )

    table_element = driver.find_element(By.CSS_SELECTOR, ".market-table__table")
    table_text = table_element.text

    lines = table_text.split('\n')

    data_rows = []
    for line in lines:
        parts = line.split()
        if len(parts) == 7:
            data_rows.append(parts)

    columns = ["Symbol", "Open", "High", "Low", "Last", "Change", "Volume"]

    df = pd.DataFrame(data_rows, columns=columns)

    def parse_price(val):
        try:
            return float(val.replace(",", ""))
        except:
            return None

    df["Last"] = df["Last"].apply(parse_price)

    high_value_df = df[df["Last"] > 1000]

    high_value_df.to_excel("psx_high_value_stocks.xlsx", index=False)
    high_value_df.to_csv("psx_high_value_stocks.csv", index=False)

    print(f"✅ Saved {len(high_value_df)} high-value stocks (>1000 PKR) to Excel and CSV.")

finally:
    driver.quit()
