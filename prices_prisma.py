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
                    return None, None, None
                link = self.specific_product_page(soup)
                link = f"https://www.prisma.fi{link}"
                print(link)
                html = self.get_html_from_url(link)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    product_name, brand_name, price = self.extract_product_information(soup, product_code)
                    print("\nTUOTE:",product_name, product_code, brand_name, price,"\n")
                    return print(product_name, link, brand_name, price)
                else:
                    print("(1) PRISMA: Error with handling HTML.")
                    return None, None, None
            else:
                print("(2) PRISMA: Error with handling HTML.")
                return None, None, None
        except Exception as e:
            return None, None, None

    def search_result_check(self, soup): #returns False if there were no search results or True if there were results
        try:
            no_search_results_element = soup.find('p', {'data-test-id': 'no-search-results', 'class': 'search_noResultTitle__PTo4Q'})
            if no_search_results_element:
                result_text = no_search_results_element.get_text(strip=True)
                return False
            else:
                return True
        except Exception as e:
            return None

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
            # Find the <div> element...
            div_elements = soup.find_all('div', class_='Accordion_accordionContent__rtLYZ accordion-content')
            # ...that contains the div...
            for div in div_elements:
                # ...that contains the <dt>-tags...
                dt_tags = div.find_all('dt')
                for dt_tag in dt_tags:
                    # ...out of which one contains the product EAN
                    if dt_tag.text.strip() == 'Tuotekoodi':
                        dd_tag = dt_tag.find_next_sibling('dd')
                        ean_value = dd_tag.text.strip()
                        print(ean_value)
                        if(int(ean_value) != product_code):
                            return ValueError("EAN code does not match prisma.fi product code")
                        break
            try:
                # Find the element with class 'ProductMainInfo_brand__NbMHp'
                brand_div = soup.find('div', class_='ProductMainInfo_brand__NbMHp')
                print(brand_div)
                # Extract the brand text from the <a> tag inside the div
                brand_name = brand_div.find('a').text
                print(brand_name)
            except Exception as e:
                # No brand name tag present in the website
                brand_name = "UNDEFINED"
            print("BRAND NAME:",brand_name)
            # Find the element with class 'ProductMainInfo_finalPrice__1hFhA'
            price_span = soup.find('span', class_='ProductMainInfo_finalPrice__1hFhA')
            print(price_span)
            # Extract the price text from the span and remove non-numeric characters
            price_text = price_span.text.replace('&nbsp;', '').replace('€', '').replace(',', '.')
            print(price_text)
            # Convert the extracted text to a float
            price = float(price_text)

            # Find the element with class 'ProductMainInfo_title__VKU68' and data-test-id 'product-name'
            product_name_h1 = soup.find('h1', class_='ProductMainInfo_title__VKU68', attrs={'data-test-id': 'product-name'})
            print(product_name_h1)
            # Extract the text content from the h1 element
            product_name = product_name_h1.text
            print(product_name)
            
            return product_name, brand_name, price
        except Exception as e:
            return None, None, None

prisma_scraper = PrismaScraper()
prisma_scraper.search_product(6418677334948)
prisma_scraper.search_product(6416129362310)



# EAN: 6418677334948
# Nimi: ABB kytkin jussi 1+1+1 10631u
# Merkki: ABB
# Hinta: 28,80

# EAN: 7314150111060 <-- ehdottaa jotain paskaa mut pitäisi olla ei tuloksia. Ja onkin :)

# EAN 982756917501 <-- Ei hakutuloksia "Hakusanalla ei löytynyt yhtään tuotetta." ku hakee prismasta

