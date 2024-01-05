# -*- coding: utf-8 -*-
"""
Created on Fri Oct 27 07:51:10 2023

@author: Kostas Panagias

Last update: 2024.01.05
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

from openpyxl import load_workbook

print('Starting Process...')
starttime = datetime.now()
print(f'Current time: {datetime.now().time().strftime("%H:%M:%S")}')

columns = ['date', 'time', 'retailer', 'product category', 'product sku','product name', 'start price', 'final price', 'availability', 'page url', 'product url','retailer sku', 'portion price']

df_scraped = pd.DataFrame(columns=columns)

# Get the current date
current_date = datetime.now()
# Format the date
formatted_date = current_date.strftime('%Y.%m.%d')

# Local files & folders

# Directory path
directory = "scrapped_results"

# Check if the directory exists
if not os.path.exists(directory):
    # Create the directory
    os.makedirs(directory)

urlinput_file = r'links\mymarket.xlsx' # This is the file (input) which contains the links of product categories for web scrapping
export_file = fr'{directory}\mymarket_{formatted_date}.xlsx' # this is the file (output) of the scrapped products

writer = pd.ExcelWriter(export_file, engine='xlsxwriter')

df = pd.read_excel(urlinput_file)

for index, row in df.iterrows():
    
    data = []
    retailer = 'Mymarket'
    product_category = row['Category']
    page_url = row['URL']
    print(f'Scraping webpage :{page_url}')
    print(f'Current Time:{datetime.now().time().strftime("%H:%M:%S")}')
    options = webdriver.ChromeOptions()
    #options.headless = True  # Run Chrome in headless mode
    options.add_argument('--headless')
    browser = webdriver.Chrome(options=options)
    browser.get(row['URL'])
    time.sleep(3)

     

    product = browser.find_elements(By.CSS_SELECTOR, 'h3')
    productname = [x.text for x in product]
    productname = productname[:-2]
    
    productlink = browser.find_elements(By.CSS_SELECTOR, 'h3 > a')
    productlink = [x.get_attribute('href') for x in productlink]
    
    productprice = browser.find_elements(By.CSS_SELECTOR, 'span[class="price"]')
    productprice = [x.text for x in productprice]
    
    productprice = [x.replace('€','') for x in productprice]
    productprice = [x.replace(',','.') for x in productprice]

    poptionprice = browser.find_elements(By.CSS_SELECTOR, 'div[class="measurment-unit-row "]')
    poptionprice = [x.text for x in poptionprice]
    poptionprice = [x.replace("\n", " - ") for x in poptionprice]
    
    retailersku = browser.find_elements(By.CSS_SELECTOR, 'div[class="sku"]')
    retailersku = [x.text for x in retailersku]
    
    productsku = ["" for x in retailersku]
    availability = ["" for x in retailersku]
    initial_price = ["" for x in retailersku]
    
    browser.quit()

    
    data = {
        'date': datetime.now().date(),
        'time': datetime.now().time().strftime("%H:%M:%S"),
        'retailer': retailer,
        'product category': product_category,
        'product sku' : productsku,
        'product name': productname,
        'start price': initial_price,
        'final price': productprice,
        'portion price': poptionprice,
        'availability': availability,
        'page url': page_url,
        'product url': productlink,
        'retailer sku': retailersku,
    }
    
    
    scrapped = pd.DataFrame(data)
    df_scraped = pd.concat([df_scraped, scrapped], ignore_index=True)
    df_scraped['final price'] = pd.to_numeric(df_scraped['final price'], errors='coerce')

df_scraped.to_excel(writer, index=False)
writer.close()


#Formatting

wb = load_workbook(export_file)

# Select the desired sheet
sheet = wb['Sheet1']  # Replace 'SheetName' with the name of your sheet

# Change the width of a specific column
# Here, 'A' is the column, and '20' is the new width
sheet.column_dimensions['F'].width = 60
sheet.column_dimensions['A'].width = 11
sheet.column_dimensions['C'].width = 13
sheet.column_dimensions['D'].width = 23
sheet.column_dimensions['M'].width = 20

# Save the workbook
wb.save(export_file)

print('Finished Process')
endtime = datetime.now()
# Calculate elapsed time
elapsed_time = endtime - starttime

# Print elapsed time in HH:MM:SS format
hours, remainder = divmod(elapsed_time.seconds, 3600)
minutes, seconds = divmod(remainder, 60)
print(f"Elapsed Time: {hours:02}:{minutes:02}:{seconds:02}")