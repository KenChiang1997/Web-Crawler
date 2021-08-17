# basic
import os
import re
import time
import numpy as np 
import pandas as pd 

# bs4
from bs4 import element 
from bs4 import BeautifulSoup

# selenium
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait


#-----------------------------------function build------------------------------------------

class funds_crawler():

    def __init__(self,selenium_path,Type):

        self.selenium_path = selenium_path
        self.df   = pd.DataFrame()

        self.Type = Type 
        if Type == "onshore" :
            self.url = "https://announce.fundclear.com.tw/MOPSonshoreFundWeb/INQ713.jsp"
        elif Type == "offshore":
            self.url = "https://announce.fundclear.com.tw/MOPSFundWeb/INQ712.jsp?commit=1&fundClassType=1&statType=1&fundAsset=all&fundInv=all&fundCurr=all&fromTo=1&dataRange=1"



    def get_data(self,desire_type,desir_page,number_of_funds):

        # chrome driver setting
        options = webdriver.ChromeOptions()  
        options.add_argument("--incognito")
        options.add_argument("--start-maximized") # Edit by Janet
        driver = webdriver.Chrome(self.selenium_path,options=options)
        driver.get(self.url)
        
        #  start fetch data
        driver.find_element_by_xpath("//input[@name ='statType' and @value=3]").click() # 點淨申購資訊
        df_data = self.df

        for i in range(1,desir_page+1): #抓多少頁ˋ
            #顯示區間
            if i >= 2 :
                driver.find_element_by_xpath("//select[@name ='fromTo']").click()
                select = Select(driver.find_element_by_xpath("//select[@name ='fromTo']"))
                select.select_by_value(str(i))

            time.sleep(1)
            btn = driver.find_element_by_xpath("//input[@name = 'btnQuery' and @value = '查詢']")
            btn.click()
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            if self.Type == "offshore" :
                df_data = df_data.append(self.offshore_clean_data(soup))
            elif self.Type == "onshore" :
                df_data = df_data.append(self.onshore_clean_data(soup))

            time.sleep(5)

        ##-----------------------把不要的基金類型刪掉-----------------------------
        df_data = df_data.loc[df_data['基金類型'] != "指數股票型"]
        df_data = df_data.loc[df_data['基金類型'] != "貨幣市場基金"]
        ##-----------------------決定要找甚麼樣的前10大---------------------------
        df_data[desire_type] = df_data.apply(lambda x : self.address_number(x[desire_type]) ,axis=1 ) # 先處理“文字(Accounting)”數字
        df_data = df_data.sort_values(by=desire_type,ascending=False) # sort value by desire type
        df_data = df_data.reset_index(drop = True)
        df_result = df_data[df_data.index < number_of_funds]
        df_result = df_result.drop(columns=["序號"])
        self.df =df_result

        return self.df


    def offshore_clean_data(self,soup):   

        i = soup.select("table")[6].text.find("單位：新台幣百萬元")
        text = soup.select("table")[6].text[i+10:]
        i = text.find("1")
        text_list = text[i:].split("\n")
        df = pd.DataFrame(columns = ["序號", "基金名稱", "基金類型", "淨申贖金額", "申購金額", "買回金額","國人持有基金", "基金規模","國人持有基金佔比","法定上限"])

        for i in range(0, len(text_list), 14):

            if len(text_list[i]) > 0 :
                df = df.append({"序號"          :  text_list[i],
                                "基金名稱"       : text_list[i+1],
                                "基金類型"       : text_list[i+2],   
                                "淨申贖金額"      : text_list[i+3],
                                "申購金額"        : text_list[i+4],
                                "買回金額"        : text_list[i+5],
                                "國人持有基金"     : text_list[i+6],
                                "基金規模"         : text_list[i+7],
                                "國人持有基金佔比"   : text_list[i+8],
                                "法定上限"          : text_list[i+9],
                }, ignore_index=True)

        return df
    
    def onshore_clean_data(self,soup):   

        i = soup.select("table")[6].text.find("單位：新台幣百萬元")
        text = soup.select("table")[6].text[i+10:]
        i = text.find("1")
        text_list = text[i:].split("\n")
        df = pd.DataFrame(columns = ["序號", "基金名稱", "基金類型", "淨申贖金額", "申購金額", "買回金額", "基金規模", "最新淨值日期","淨值"])

        for i in range(0, len(text_list), 13):

            if len(text_list[i]) > 0 :
                df = df.append({"序號"     : text_list[i],
                                "基金名稱"  : text_list[i+1],
                                "基金類型"   : text_list[i+2],
                                "淨申贖金額"  : text_list[i+3],
                                "申購金額"    : text_list[i+4],
                                "買回金額"     : text_list[i+5],
                                "基金規模"      : text_list[i+6],
                                "最新淨值日期"   : text_list[i+7],
                                "淨值"          : text_list[i+8],
                }, ignore_index=True)

        return df


    def address_number(self,number):
        number = str(number).replace(",","")
        return float(number)





if __name__ == '__main__':

    crawler    = funds_crawler(selenium_path="/Users/chen-lichiang/Desktop/chromedriver",Type='onshore')
    crawler_df = crawler.get_data(desire_type="淨申贖金額",desir_page=3,number_of_funds=10)
    print(crawler_df)


