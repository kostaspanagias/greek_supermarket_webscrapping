# -*- coding: utf-8 -*-
"""
Created on Thu Oct 12 08:36:09 2023

@author: kpanagias
"""


import pandas as pd

import os
from datetime import datetime
import time
import math

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service



from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

columns = ['date', 'time', 'retailer', 'product category', 'product sku','product name', 'start price', 'final price', 'availability', 'page url', 'product url','retailer sku', 'portion price']

df_scraped = pd.DataFrame(columns=columns)

# Get the current date
current_date = datetime.now()
# Format the date
formatted_date = current_date.strftime('%Y.%m.%d')


# Local files & folders
linksfolder = r'C:\SuperMarket prices\links'
sklavenitisfile = r'C:\SuperMarket prices\links\sklavenitis.xlsx' # This is the file (input) which contains the links of product categories for web scrapping
export_file = fr'C:\SuperMarket prices\scrapped_results_sklavenitis_{formatted_date}.xlsx' # this is the file (output) of the scrapped products


writer = pd.ExcelWriter(export_file, engine='xlsxwriter')

df = pd.read_excel(sklavenitisfile)

for index, row in df.iterrows():
    
    data = []
    retailer = 'Sklavenitis'
    product_category = row['Category']
    page_url = row['URL']
    print(f'Scraping webpage :{page_url}')
    options = webdriver.ChromeOptions()
    #options.headless = True  # Run Chrome in headless mode
    options.add_argument('--headless')
    browser = webdriver.Chrome(options=options)
    browser.get(row['URL'])
    time.sleep(3)

     

    product = browser.find_elements(By.CSS_SELECTOR, 'h4[class="product__title"]')
    productname = [x.text for x in product]
    
    productlink = browser.find_elements(By.CSS_SELECTOR, 'h4 > a')
    productlink = [x.get_attribute('href') for x in productlink]
    
    productprice = browser.find_elements(By.CSS_SELECTOR, 'div[class="price"]')
    productprice = [x.text for x in productprice]
    
    productprice = [x[:x.find('€/')].strip() for x in productprice]
    productprice = [x.replace(',','.') for x in productprice]

    poptionprice = browser.find_elements(By.CSS_SELECTOR, 'div[class="priceKil"]')
    poptionprice = [x.text for x in poptionprice]
    poptionprice = [x.replace("\n", " - ") for x in poptionprice]
    
    productsku = browser.find_elements(By.CSS_SELECTOR, 'span[class="dropdown__trigger icon-favlist add-to-list"]')
    productsku = [x.get_attribute('data-sku') for x in productsku]
    
    browser.quit()

    
    data = {
        'date': datetime.now().date(),
        'time': datetime.now().time(),
        'retailer': retailer,
        'product category': product_category,
        'product sku' : productsku,
        'product name': productname,
        #'start price': initial_price,
        'final price': productprice,
        'portion price': poptionprice,
        #'availability': availability,
        'page url': page_url,
        'product url': productlink,
        'retailer sku': productsku,
    }
    
    
    scrapped = pd.DataFrame(data)
    df_scraped = df_scraped.append(scrapped, ignore_index=True)
    df_scraped['final price'] = pd.to_numeric(df_scraped['final price'], errors='coerce')

df_scraped.to_excel(writer, index=False)
writer.save()
writer.close()
