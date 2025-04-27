import requests
from bs4 import BeautifulSoup
def extract_books_and_prices(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        book_data = []
        book_links = soup.find_all('a', class_='full-unstyled-link')
        for link in book_links:
            book_info = link.get_text(strip=True)
            if book_info and "by" in book_info:
                price_tag = link.find_parent('div').find('span', class_='price')  # Adjust class name for price if needed
                price = price_tag.get_text(strip=True) if price_tag else 'Price not available'

                book_data.append({'book_info': book_info, 'price': price})

        if book_data:
            print(f"\nFound {len(book_data)} books:")
            for idx, data in enumerate(book_data, 1):
                print(f"{idx}. {data['book_info']} - {data['price']}")
        else:
            print("No books found.")

    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
def search_category_and_scrape(url, category):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
        }

        search_url = f"{url}/search?q={category.replace(' ', '+')}"
        response = requests.get(search_url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        category_data = []
        category_links = soup.find_all('a', class_='full-unstyled-link')

        for link in category_links:
            product_name = link.get_text(strip=True)
            product_url = link.get('href', '')
            if product_name and product_url:
                category_data.append({'product_name': product_name, 'product_url': f"{url}{product_url}"})

        if category_data:
            print(f"\nFound {len(category_data)} products in '{category}':")
            for idx, data in enumerate(category_data, 1):
                print(f"{idx}. {data['product_name']} - Link: {data['product_url']}")
        else:
            print(f"No products found for category '{category}'.")

    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    target_url = input("Enter URL to scrape: ")
    extract_books_and_prices(target_url)

    category_name = input("\nEnter the category name to search and scrape: ")
    search_category_and_scrape(target_url, category_name)
