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


class PrismaScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        try:
            url = f"https://www.prisma.fi/haku?search={product_code}"
            print(url)
            html = self.get_html_from_url(url)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                if not self.search_result_check(soup):
                    raise ValueError("\nPRISMA: Tuotesivua ei löytynyt")
                link = self.specific_product_page(soup)
                link = f"https://www.prisma.fi{link}"
                print(link)
                html = self.get_html_from_url(link)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    product_name, brand_name, price = self.extract_product_information(soup, product_code)
                    print("\nPRISMA:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                    return product_name, link, brand_name, price
                else:
                    raise ValueError("(1) PRISMA: Error with handling HTML.")
            else:
                raise ValueError("(2) PRISMA: Error with handling HTML.")
        except Exception as e:
            return None, None, None, None

    def search_result_check(self, soup): #returns False if there were no search results or True if there were results
        try:
            no_search_results_element = soup.find('p', {'data-test-id': 'no-search-results', 'class': 'search_noResultTitle__PTo4Q'})
            if no_search_results_element:
                #result_text = no_search_results_element.get_text(strip=True)
                return False
            else:
                return True
        except Exception as e:
            return False

    def specific_product_page(self, soup):
        try:
            link_tag = soup.find('a', class_='ProductCard_blockLink__NTUGB') # Find the <a> tag with the specified class
            if link_tag:
                link = link_tag['href'] # Extract the value of the href attribute
                return link
            else:
                return None
        except Exception as e:
            return None
        
    def extract_product_information(self, soup, product_code):
        try:
            # Initialize the WebDriver for interaction
            url = self.driver.current_url
            self.driver.get(url)
            
            # Wait for the accordion button to be present
            accordion_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-test-id="accordion-trigger-product-feature-item"]'))
            )
            accordion_button.click()

            # Wait for the content to be present after clicking the button
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[data-test-id="accordion-content-product-feature-item"]'))
            )

            # Get the updated page source and parse it with BeautifulSoup
            html = self.driver.page_source
            soup = BeautifulSoup(html, 'html.parser')

            # Extract the EAN number
            ean_number = None
            dt_elements = soup.find_all('dt')
            for dt in dt_elements:
                if 'Tuotekoodi' in dt.text:
                    ean_number = dt.find_next_sibling('dd').text.strip()
                    break
            if int(ean_number) != product_code:
                raise ValueError("PRISMA: EAN-koodi ei vastannut hakua")
            
            # Extract the product name from the specified h1 tag
            product_name_h1 = soup.find('h1', class_='text-heading-small-regular pb-xxsmall sm:pb-medium', attrs={'data-test-id': 'product-name'})
            product_name = product_name_h1.text.strip()
            
            brand_name="??"
            brand_div = soup.find('div', class_='text-body-medium-regular mb-xxsmall text-color-text-stronger-neutral')
            if brand_div:
                brand_name = brand_div.find('a', {'data-test-id': 'product-brand-link'}).text

            price_span = soup.find('span', class_='text-heading-small-medium flex sm:text-heading-medium-medium')
            price_text = price_span.text.strip()
            # Remove the euro sign and non-breaking space
            price = price_text.replace('€', '').replace('\xa0', '')
            return product_name, brand_name, price

        except Exception as e:
            print(f"Error extracting product information: {str(e)}")
            return None, None, None

prisma_scraper = PrismaScraper()
#prisma_scraper.search_product(6418677334948)
#prisma_scraper.search_product(7314150111060)
#prisma_scraper.search_product(982756917501)
prisma_scraper.search_product(6416129362310)



# EAN: 6418677334948
# Nimi: ABB kytkin jussi 1+1+1 10631u
# Merkki: ABB
# Hinta: 28,80

# EAN: 7314150111060 <-- ehdottaa jotain paskaa mut pitäisi olla ei tuloksia. Ja onkin :)

# EAN 982756917501 <-- Ei hakutuloksia "Hakusanalla ei löytynyt yhtään tuotetta." ku hakee prismasta

