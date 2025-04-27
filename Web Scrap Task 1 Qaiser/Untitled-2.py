import requests
from bs4 import BeautifulSoup

def extract_book_titles(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
        }

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')

        possible_tags = ['h1', 'h2', 'h3', 'h4', 'a', 'span', 'p', 'div']
        seen_titles = set()
        book_titles = []

        for tag in possible_tags:
            elements = soup.find_all(tag)
            for el in elements:
                text = el.get_text(strip=True)
                if text and 3 < len(text) < 100 and text.lower() != 'book':
                    if text not in seen_titles:
                        seen_titles.add(text)
                        book_titles.append(text)

        if book_titles:
            print(f"\nFound {len(book_titles)} possible book titles on the first page:")
            for idx, title in enumerate(book_titles, 1):
                print(f"{idx}. {title}")
        else:
            print("No book titles found on the first page.")

    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    target_url = input("Enter URL to scrape: ")
    extract_book_titles(target_url)
