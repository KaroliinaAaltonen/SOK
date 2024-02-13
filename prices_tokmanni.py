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
            url = f"https://www.tokmanni.fi/search/?q={product_code}"
            html = self.get_html_from_url(url, 10)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                self.search_result_check(soup)
                link = self.specific_product_page(soup, product_code)
                html = self.get_html_from_url(link, 10)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    price, product_name = self.extract_product_information(soup, product_code)
                    if product_name and price:
                        print("\nTOKMANNI:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                        return link, price, product_name
                else:
                    ValueError("(1) TOKMANNI: Error with handling HTML.")
            else:
                ValueError("(2) TOKMANNI: Error with handling HTML.")
        except Exception as e:
            return None, None, None

    def search_result_check(self, soup): # returns False if no results found, True if results were found, otherwise None
        try:
            # Define the target text
            target_text = "Kokeile toista hakusanaa ..."
            page_wrapper = soup.find('div', class_='page-wrapper')
            #print("PAGE_WRAPPER:",page_wrapper)
            page_columns = page_wrapper.find('div', class_='columns')
            #print("COLUMNS:",page_columns)
            main_column = page_columns.find('div', class_='column main')
            #print("MAIN COLUMN:",main_column)
            no_results_parent = main_column.find('div', class_='kuContainer kuFiltersLeft')
            #print("PARENT:",no_results_parent)
            no_results = no_results_parent.find('div', class_='kuNoRecordFound')
            #print("CHILD:",no_results)
            child_1 = no_results.find('div', class_='kuNoResults-lp')
            #print("CHILD1:",child_1)
            child_2 = child_1.find('div', class_='kuNoResultsInner')
            #print("CHILD2:",child_2)
            target_div = child_2.find('div', class_='kuNoResults-lp-message')
            #print("TARGET:",target_div)
            no_results_text = target_div.get_text()
            #print("TEXT:",no_results_text)

            # Check if the target text is present in the div
            if target_text in no_results_text:
                return ValueError("TOKMANNI: tuotetta ei löytynyt hausta")
            else:
                return True
        except Exception as e:
            return ValueError("TOKMANNI: tuotetta ei löytynyt hausta")

    def specific_product_page(self, soup, product_code): # Returns the link to the product page ro None
        try:
            # Define the target text
            target_text = "Kokeile toista hakusanaa ..."
            page_wrapper = soup.find('div', class_='page-wrapper')
            #print("PAGE_WRAPPER:",page_wrapper)
            page_columns = page_wrapper.find('div', class_='columns')
            #print("COLUMNS:",page_columns)
            main_column = page_columns.find('div', class_='column main')
            #print("MAIN COLUMN:",main_column)
            results_parent = main_column.find('div', class_='kuContainer kuFiltersLeft')
            #print("PARENT:",results_parent)
            results = results_parent.find('div', class_='kuProListing')
            #print("CHILD:",results)
            child_1 = results.find('div', class_='kuResultList')
            #print("CHILD1:",child_1)
            child_2 = child_1.find('div', class_='kuProductContent')
            #print("CHILD2:",child_2)
            a_tag = child_2.find('div', class_='kuGridView').find('a')
            #print("TARGET:",a_tag)
            link = a_tag['href']
            #print(url)

            # Check if the link was found
            if link:
                return link
            else:
                return ValueError("TOKMANNI: tuotesivua ei löytynyt.")
        except Exception as e:
            return ValueError("TOKMANNI: tuotesivua ei löytynyt.")
        
    def extract_product_information(self, soup, product_code): # Returns price and product name or None
        try:
            ean_div = soup.find('div', class_='product attribute sku')
            ean_value = ean_div.find('div', class_='value').get_text(strip=True)
            if(str(ean_value) != str(product_code)):
                return ValueError("Tokamnni: Tuotesivun EAN ei täsmää haettua tuotetta.")
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
            return ValueError("Tokmanni: Virhe tuotetietojen hakemisessa.")

tokmanni_scraper = TokmanniScraper()
tokmanni_scraper.search_product(6438140053527)
tokmanni_scraper.search_product(6418677334962)
tokmanni_scraper.search_product(982756917501)


# EAN: 6438140053527 <-- sivu ehdottaa jotain paskaa, pitäis olla ei tuloksia

# EAN: 6418677334962
# Nimi: Kytkin ABB Jussi 7 1067UP
# Merkki: ABB
# Hinta: 21,99

# EAN 982756917501 <-- Ei hakutuloksia

