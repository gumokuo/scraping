#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 14 11:15:45 2019

@author: kenkuo
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import xlwings as xw

wb=xw.Book('日本語文法まとめ.xlsx')
sht=wb.sheets('まとめ')

# chrome_options = Options()
# chrome_options.add_argument('--headless')
driver_path = "/Users/kenkuo/Downloads/chrome-mac-x64/Google Chrome for Testing.app"
# driver = webdriver.Chrome(driver_path,options=chrome_options)
# driver = webdriver.Chrome(driver_path)
driver = webdriver.Chrome()
driver.get("https://nihongonosensei.net/?page_id=10246#linkn1")
# driver.get("https://nihongonosensei.net/?p=20509")


# list_a_href=[]
shift=1
for i in range(52,54): #171
    # a_href=driver.find_element(By.XPATH,'//*[@id="mouseover1"]/tbody/tr['+str(i)+']/td[3]/a').click()
    a_href=driver.find_element(By.XPATH,'/html/body/div[1]/div/main/article/div/div/div[1]/table[2]/tbody/tr['+str(i+2)+']/td/a').click()
    # list_a_href.append(a_href.get_attribute('href'))
    search = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'midashi')))
    
    txt=driver.find_elements(By.XPATH,'//*[@id="mainEntity"]/div[1]/p')
    cnt=0
    is_bikou_pos=False
    for j in txt:
        # print(i.text)
        if j.text=="接続":
            setsuzoku_pos=cnt
            # print(setsuzoku_pos)
            # sht.cells(2,5).value=plist[cnt+1].text
        if j.text=="意味":
            imi_pos=cnt
            t=''
            for k in range(setsuzoku_pos+1,imi_pos):
                t=t+txt[k].text
            # print(imi_pos)
            # print(t)
            sht.cells(i+shift,5).value=t
        if j.text=="解説":
            kaisetsu_pos=cnt
            t=''
            for k in range(imi_pos+1,kaisetsu_pos):
                t=t+txt[k].text
            # print(kaisetsu_pos)
            # print(t)
            sht.cells(i+shift,6).value=t
        if j.text=="例文":
            reibunn_pos=cnt
            t=''
            for k in range(kaisetsu_pos+1,reibunn_pos):
                t=t+txt[k].text
            # print(reibunn_pos)
            # print(t)
            sht.cells(i+shift,7).value=t
        if j.text=="備考":
            is_bikou_pos=True
            bikou_pos=cnt
            t=''
            tt=''
            for k in range(reibunn_pos+1,bikou_pos):
                # if "（１）" in txt[cnt+1].text:
                t=t+txt[k].text
            # print(bikou_pos)
            sht.cells(i+shift,8).value=t
            # print(txt[cnt+1].text)
        if is_bikou_pos:
            if cnt>bikou_pos:
                tt=tt+txt[cnt].text
        cnt+=1
    sht.cells(i+shift,9).value=tt
    driver.back()
    time.sleep(10)
    # search = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'linkn0')))