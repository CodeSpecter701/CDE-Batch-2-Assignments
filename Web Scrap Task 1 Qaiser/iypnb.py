import requests
from bs4 import BeautifulSoup
url = 'https://www.thriftbooks.com/browse/?b.search=python'
headers = {'User-Agent': 'Mozilla/5.0'}
response = requests.get(url, headers=headers)

soup = BeautifulSoup(response.text, 'html.parser')
book_titles = soup.find_all('a', class_='bookTitle')

for i, tag in enumerate(book_titles, 1):
    title = tag.get_text(strip=True)
    print(f"{i}. {title}")
