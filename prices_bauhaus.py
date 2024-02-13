import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
from openpyxl.utils import get_column_letter
import winsound

################################################################
#                                                              #
#                                                              #
# HOURS WASTED WITH THIS: 10                                   #
#increment as you go                                           #
#                                                              #
#                                                              #
#                                                              #
################################################################
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


class BauhausScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        try:
            url = f"https://www.bauhaus.fi/catalogsearch/result/?q={product_code}"
            html = self.get_html_from_url(url, 10)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                self.search_result_check(soup)
                link = self.specific_product_page(soup, product_code)
                if not link:
                    return None, None, None
                html = self.get_html_from_url(link, 10)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    price, product_name = self.extract_product_information(soup)
                    if product_name and price:
                        #print("\nBAUHAUS:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                        return link, price, product_name
                else:
                    print("(1) BAUHAUS: Error with handling HTML.")
                    return None, None, None
            else:
                print("(2) BAUHAUS: Error with handling HTML.")
                return None, None, None
        except Exception as e:
            return None, None, None

    def search_result_check(self, soup): # returns False if no results found, True if results were found, otherwise None
        try:
            # Display the entire HTML content
            print(soup.prettify())

        except Exception as e:
            print(e)

    def specific_product_page(self, soup, product_code): # Returns the link to the product page ro None
        try:
            # Find the <em> tag
            em_tag = soup.find('em')
            # Check if the text inside the <em> tag is equal to product_code
            if em_tag and int(em_tag.text.strip()) == int(product_code):
                # Find the <a> tag and get the href value
                a_tag = soup.find('a', {'jsname': 'UWckNb'})
                if a_tag:
                    link = a_tag['href']
                    if "bauhaus" in link:
                        return link
        except Exception as e:
            return None
        
    def extract_product_information(self, soup): # Returns price and product name or None
        try:
            # Find the <span> element with class "price"
            price_span = soup.find('span', class_='price')
            # Extract the text content from the price_span
            price_text = price_span.get_text(strip=True) 
            price = float(price_text.replace('\xa0', ''))/100

            # Find the <span> element with class "base" and itemprop "name"
            name_span = soup.find('h1', class_='page-title').find('span', class_='base')
            # Extract the text content from the name_span
            product_name = name_span.get_text(strip=True) 

            return price, product_name
        except Exception as e:
            return None, None

bauhaus_scraper = BauhausScraper()
bauhaus_scraper.search_product(6418677334948)
bauhaus_scraper.search_product(7314150111060)
bauhaus_scraper.search_product(982756917501)



# EAN: 6418677334948
# Nimi: ABB kytkin jussi 1+1+1 10631u
# Merkki: ABB
# Hinta: 28,80

# Vitut tästä
# EAN: 7314150111060 <-- ehdottaa jotain paskaa mut pitäisi olla ei tuloksia. Ja onkin :)

# EAN 982756917501 <-- Ei hakutuloksia "Hakusanalla ei löytynyt yhtään tuotetta." ku hakee prismasta

