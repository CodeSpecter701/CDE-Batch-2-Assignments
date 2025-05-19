from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import sys

driver_path = 'D:/Selenium Project Work!/chromedriver.exe'
service = Service(driver_path)
options = Options()

driver = webdriver.Chrome(service=service, options=options)

try:
    driver.get("https://www.psx.com.pk/market-summary/")

    # Wait for table rows instead of container
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

    df.to_excel("psx_all_data_snapshot.xlsx", index=False)
    df.to_csv("psx_all_data_snapshot.csv", index=False)

    print(f"✅ Saved {len(df)} rows of data to Excel and CSV.")

finally:
    driver.quit()
