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


class TokmanniScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        try:
            query = f"\"tokmanni\" \"{product_code}\""
            url = f"https://www.google.com/search?q={query}"
            html = self.get_html_from_url(url, 10)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                if not self.search_result_check(soup):
                    #print("\nTOKMANNI:\nEi tuloksia haulla:", product_code)
                    return None, None, None
                link = self.specific_product_page(soup, product_code)
                if not link:
                    return None, None, None
                html = self.get_html_from_url(link, 10)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    price, product_name = self.extract_product_information(soup)
                    if product_name and price:
                        #print("\nTOKMANNI:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                        return print(link, price, product_name)
                else:
                    print("(1) TOKMANNI: Error with handling HTML.")
                    return None, None, None
            else:
                print("(2) TOKMANNI: Error with handling HTML.")
                return None, None, None
        except Exception as e:
            return None, None, None

    def search_result_check(self, soup): # returns False if no results found, True if results were found, otherwise None
        try:
            # Define the target text
            target_text = "About 0 results"

            # Find the div element with class "LHJvCe" and id "result-stats"
            result_stats_div = soup.find('div', {'class': 'LHJvCe', 'id': 'result-stats'})

            # Check if the target text is present in the div
            if result_stats_div and target_text in result_stats_div.get_text():
                return False
            else:
                return True
        except Exception as e:
            return None

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
                    return link
        except Exception as e:
            return None
        
    def extract_product_information(self, soup): # Returns price and product name or None
        try:
            # Find the span element with class "price-wrapper"
            price_wrapper_span = soup.find('span', class_='price-wrapper')
            # Extract the value from the data-price-amount attribute
            price = price_wrapper_span.get('data-price-amount')
            
            # Find the h1 element with class "page-title"
            page_title_h1 = soup.find('h1', {'class': 'page-title'})
            # Extract product name
            product_name = page_title_h1.get_text(strip=True)
            return price, product_name
        except Exception as e:
            return None, None

tokmanni_scraper = TokmanniScraper()
tokmanni_scraper.search_product(6418677334948)
tokmanni_scraper.search_product(7314150111060)
tokmanni_scraper.search_product(982756917501)



# EAN: 6418677334948
# Nimi: Kytkin ABB Jussi 1+1+1 10631UP
# Merkki: ABB
# Hinta: 22,95

# EAN: 7314150111060
# Nimi: Hylsysarja Bahco S460 46-osainen
# Merkki: Bahco
# Hinta: 99,00

# EAN 982756917501 <-- Ei hakutuloksia

