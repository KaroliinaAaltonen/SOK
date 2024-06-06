import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
from openpyxl.utils import get_column_letter
import winsound

# web scraping functions
# driver is started locally (and ended)
# get_html_from_url() uses sleep=10 with puuilo.fi and tokmanni.fi as per their robots.txt
class Scraper:
    def __init__(self):
        self.driver = self.local_driver()

    def local_driver(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # Run Chrome in headless mode (without UI)
        service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver

    def end_session(self):
        self.driver.quit()

    def get_html_from_url(self, url, sleep=1):
        try:
            self.driver.get(url)
            time.sleep(sleep)
            return self.driver.page_source
        except Exception as e:
            print(f"Error getting HTML from {url}: {str(e)}")
            return None


class KRautaScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        try:
            url = f"https://www.k-rauta.fi/etsi?query={product_code}"
            html = self.get_html_from_url(url)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                if not self.search_result_check(soup):
                    raise ValueError("\nK-RAUTA:\nEi tuloksia haulla:", product_code)
                
                link = self.specific_product_page(soup)
                link = f"https://www.k-rauta.fi{link}"
                html = self.get_html_from_url(link)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    price, product_name = self.extract_product_information(soup)
                    if product_name and price:
                        print("\nK-RAUTA:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                        return link, price, product_name
                else:
                    raise ValueError("(1) K-RAUTA: Error with handling HTML.")
            else:
                raise ValueError("(2) K-RAUTA: Error with handling HTML.")
        except Exception as e:
            return None, None, None

    def search_result_check(self, soup): # returns True if there are search results or False if no results
        try:
            no_search_results_element = soup.find('div', class_="empty-search-result")
            if no_search_results_element:
                #result_text = no_search_results_element.get_text(strip=True)
                return False
            else:
                return True
        except Exception as e:
            return None
        
    def specific_product_page(self, soup): # returns specific product page or None
        try:
            a_element = soup.find('a', class_='product-card') # Find the <a> element with class "product-card"
            if a_element:
                link = a_element['href'] # Extract the value of the "href" attribute
                return link
            else:
                return None
        except Exception as e:
            return None

    def extract_product_information(self, soup): # returns price and product name or None
        try:
            # Extract price span
            price_span = soup.find('span', class_='price-view-fi__sale-price--prefix')
            price_text = price_span.next_sibling.strip()
            price= float(price_text.replace("€", "").replace(",", "."))

            # Extract the product name
            product_name = soup.find('h1', class_='product-heading__product-name').text.strip()
            return price, product_name
        except Exception as e:
            return None, None

krauta_scraper = KRautaScraper()
krauta_scraper.search_product(6418677334948)
krauta_scraper.search_product(7314150111060)
krauta_scraper.search_product(982756917501)



# EAN: 6418677334948
# Nimi: ABB kytkin jussi 1+1+1 10631u
# Merkki: ABB
# Hinta: 28,80

# EAN: 7314150111060 <-- Ei hakutuloksia. "Valitettavasti emme löytäneet mitään hakusanalla"

# EAN 982756917501 <-- Ei hakutuloksia "Valitettavasti emme löytäneet mitään hakusanalla..."

