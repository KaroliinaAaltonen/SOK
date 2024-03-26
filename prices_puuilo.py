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


class PuuiloScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        try:
            url = f"https://www.puuilo.fi/catalogsearch/result/?q={product_code}"
            html = self.get_html_from_url(url)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                self.search_result_check(soup)
                link = self.specific_product_page(soup)
                html = self.get_html_from_url(link)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    price, product_name = self.extract_product_information(soup, product_code)
                    if product_name and price:
                        return link, price, product_name
                else:
                    return ValueError("(1) PUUILO: Error with handling HTML.")
            else:
                return ValueError("(2) PUUILO: Error with handling HTML.")
        except Exception as e:
            return None, None, None
    
    def search_result_check(self, soup): # returns True if there are search results or False if no results
        hits_div = soup.find('div', class_='hits')
        result_count = hits_div.find('span', class_='ais-Stats-text')
        if result_count and result_count.get_text(strip=True) == '0 hakutulosta':
            return False
        else:
            return True
        
    def specific_product_page(self, soup): # returns link to the prodcut page or None
        try:
            div_list_content = soup.find('div', class_='list-content')
            div_result_wrapper = div_list_content.find('div', class_='result-wrapper')
            a_element = div_result_wrapper.find('a', class_='result') # Find the <a> element inside the parent div tag
            if a_element:
                link = a_element['href'] # Extract the value of the "href" attribute
                return link
            else:
                return ValueError("PUUILO: Tuotesivua ei löytynyt.")
        except Exception as e:
            return None

    def extract_product_information(self, soup, product_code): # returns price and product name or None
        try:
            # Find the <div> where EAN code is nested
            ean_parent_div = soup.find('div', class_='product-info-main')
            ean_child_div = ean_parent_div.find('div', class_='product-sku-container')
            ean_content_div = ean_child_div.find('div', class_='product attribute ean')
            ean = ean_content_div.find('div', class_='value').get_text()
            if(int(ean) != product_code):
                return ValueError("PUUILO: EAN-koodi ei vastannut hakua")
            # Find the element with class 'price' inside the span with id 'product-price-107111'
            price_span = soup.find('span', class_='price')
            # Extract the price text from the span and remove non-numeric characters
            price_text = price_span.text.replace('&nbsp;', '').replace('€', '').replace(',', '.')
            # Convert the extracted text to a float
            price = float(price_text)

            # Extract product name
            product_name = soup.find('span', {'data-ui-id': 'page-title-wrapper'}).text.strip()

            return price, product_name
        except Exception as e:
            return None, None

puuilo_scraper = PuuiloScraper()
puuilo_scraper.search_product(6418677334948)
puuilo_scraper.search_product(7314150111060)


# EAN: 6418677334948
# Nimi: Jussi kytkin 1+1+1 uppoasennus
# Merkki: -
# Hinta: 21,90

# EAN: 7314150111060 <-- Ei tuloksia
