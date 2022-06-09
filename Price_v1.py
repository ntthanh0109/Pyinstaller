from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from time import sleep
import pandas as pd
import numpy as np
import csv
import os
import ctypes
from selenium.common import exceptions
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from webdriver_manager.chrome import ChromeDriverManager
driver = webdriver.Chrome(ChromeDriverManager().install())
url = 'https://www.ashford.com/'
driver.get(url)
ctypes.windll.user32.MessageBoxW(
    None, "Choose exel file to input data:", "Notice", 0x40000)
Tk().withdraw()
filename = askopenfilename()
data = pd.read_excel(filename)
df = pd.DataFrame(data, columns=['SKU'])
np.savetxt('sku.txt', df.values, fmt='%s')
sku = open('sku.txt')
line = sku.readlines()


def GetIn4():
    page_source = BeautifulSoup(driver.page_source, "html.parser")
    info_div = page_source.find('div', {'class': 'product-info-price'})
    info_div_1 = page_source.find(
        'div', {'class': 'd-flex align-items-end justify-content-flex-start text-red mt-3'})
    try:
        name = page_source.find(
            'a', {'class': 'f-17 qvPrdURL link-black text-capitalize'}).get_text().strip()
        o_price = info_div.find('span', class_='price').get_text().strip()
        sale_off = page_source.find(
            'div', class_='price-percentage clear').get_text().strip()
        n_price = info_div_1.find('span', class_='price').get_text().strip()
        writer.writerow({headers[0]: name, headers[1]: i, headers[2]: o_price, headers[3]: sale_off, headers[4]: n_price})
        print('\n')
    except:
        pass


ctypes.windll.user32.MessageBoxW(
    None, "Choose csv file to save output:", "Notice", 0x40000)
Tk().withdraw()
fileout = askopenfilename()
with open(fileout, 'w',  newline='') as file_output:
    headers = ['Name', 'SKU', 'Old Price', 'Sale', 'New Price']
    writer = csv.DictWriter(file_output, delimiter=',',
                            lineterminator='\n', fieldnames=headers)
    writer.writeheader()
    for i in line:
        try:
            search_field = driver.find_element_by_id('search')
            search_field.send_keys(i + Keys.ENTER)
            sleep(2)
            GetIn4()
        except exceptions.StaleElementReferenceException:
            pass
ctypes.windll.user32.MessageBoxW(
    None, "Open result", "COMPLETE UPDATE", 0x40000)
os.system("start " + fileout)
