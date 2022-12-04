from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
from PIL import Image
import pandas as pd 
from selenium.common.exceptions import NoSuchElementException
import os 
from shutil import rmtree
from selenium.webdriver.common.keys import Keys 

from bs4 import BeautifulSoup
import requests
import re

# Here Chrome  will be used
driver = webdriver.Chrome(executable_path= '/home/laxmi/Documents/Kaggle_Project/chromedriver')
 
# URL of website
url = "https://www.nepalstock.com.np/floor-sheet"
 
# # Opening the website
driver.get(url)
time.sleep(10)

element = driver.find_element('xpath', "/html/body/app-root/div/main/div/app-floor-sheet/div/div[3]/div/div[5]/div/select")
driver.execute_script("arguments[0].click();", element)
time.sleep(2)

def total_turnover():
    total_turnover = driver.find_element(by= By.XPATH, value= '/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[1]/table/tbody/tr/td[1]').text
    return total_turnover
def data_collection():
    # Items Per Page
    
    try:
        select = Select( driver.find_element(by=By.XPATH, value="//select[@class='ng-untouched ng-pristine ng-valid']"))
        select.select_by_visible_text('500')
    except:
        pass 
    driver.find_element(by=By.XPATH, value=" //button[normalize-space()='Filter']").click()
    time.sleep(1) 
    end_of_page = driver.find_element(by= By.XPATH, value='/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[9]/a/span[2]').text
    print(end_of_page)
    details = []
    for j in range(int(end_of_page)):
        try:
            
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            all_tables_value = soup.findAll('td')

            time.sleep(1)

            all_tables_value = soup.findAll('td')
            all_details = []
            for a in all_tables_value:
                all_details.append(a.text)
            
            for i in range(0, len(all_details), 8):
                try:
                    SN = str(all_details[i])
                    contract_no = str(all_details[i+1]).strip()
                    stock_symbol = str(all_details[i+2]).strip()
                    Buyer_Number = str(all_details[i+3]).strip()
                    Seller_Number = str(all_details[i+4]).strip()
                    Quantity = str(all_details[i+5]).replace(',', '').replace('"', ' ').strip()
                    Rate = str(all_details[i+6]).replace(',', '').replace('"', ' ').strip()
                    Amount = str(all_details[i+7]).replace(',', '').replace('"', ' ').strip()
                    detail = [SN, contract_no, stock_symbol, Buyer_Number, Seller_Number, Quantity, Rate, Amount]
                    details.append(detail)
                except:
                    pass
            element = driver.find_element('xpath', '/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[2]/pagination-controls/pagination-template/ul/li[10]/a')
            driver.execute_script("arguments[0].click();", element) 
            time.sleep(3)
            print(j)
        except: 
            print('Complete all files')
            df1 = pd.DataFrame(details, columns= ['S.N.', 'Contract No', 'Stock Symbol', 'Buyer Broker', 'Seller Broker', 'Quantity', 'Rate', 'Amount'])
            df1.to_csv('buy_sell_broker_except.csv'.format())

    df = pd.DataFrame(details, columns= ['S.N.', 'Contract No', 'Stock Symbol', 'Buyer Broker', 'Seller Broker', 'Quantity', 'Rate', 'Amount'])
    df.to_csv('buy_sell_broker_details.csv'.format())



total_quantity = driver.find_element(by= By.XPATH, value='/html/body/app-root/div/main/div/app-floor-sheet/div/div[5]/div[1]/table/tbody/tr/td[2]').text
print(total_quantity)
