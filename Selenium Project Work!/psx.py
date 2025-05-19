from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
import pandas as pd
import time

# Set path to ChromeDriver
driver_path = 'D:/Selenium Project Work!/chromedriver.exe'
service = Service(driver_path)
options = Options()
# options.add_argument('--headless')  # Uncomment to run without opening browser

# Start ChromeDriver
driver = webdriver.Chrome(service=service, options=options)

# Load the PSX Market Summary page
driver.get("https://www.psx.com.pk/market-summary/")

# Wait until at least one table is loaded
WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.TAG_NAME, "table"))
)

# Give extra time for full rendering
time.sleep(3)

# Get the full page source after JavaScript loads content
html = driver.page_source
driver.quit()

# Parse the HTML with BeautifulSoup
soup = BeautifulSoup(html, "html.parser")

# Find all tables
tables = soup.find_all("table")
if not tables:
    print("❌ No tables found.")
    exit()

# Pick the first table
table = tables[0]

# Extract headers
headers = [th.get_text(strip=True) for th in table.find_all("th")]

# Extract all rows
data = []
for row in table.find_all("tr")[1:]:  # skip header row
    cells = row.find_all("td")
    if cells:
        row_data = [cell.get_text(strip=True) for cell in cells]
        data.append(row_data)

# Create DataFrame and save to Excel
df = pd.DataFrame(data, columns=headers)
df.to_excel("psx.xlsx", index=False)
print("✅ PSX data saved successfully to psx.xlsx")
