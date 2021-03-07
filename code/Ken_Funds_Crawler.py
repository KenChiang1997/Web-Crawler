import pandas as pd
from time import sleep
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
#-----------------------------------function build------------------------------------------
class funds_crawler():
    def __init__(self,selenium_path):
        self.selenium_path = selenium_path
        self.df = pd.DataFrame()
    def get_data(self,desire_type,desir_page,number_of_funds):
        url = "https://announce.fundclear.com.tw/MOPSonshoreFundWeb/INQ713.jsp"
        driver = webdriver.Chrome(self.selenium_path)
        driver.get(url)
        attribute = driver.find_element_by_xpath("//input[@name ='statType' and @value=3]")
        attribute.click()
        df_data=self.df
        for i in range(1,desir_page): #抓多少頁ˋ
            #顯示區間
            if i >= 2:
                driver.find_element_by_xpath("//select[@name ='fromTo']").click()
                select = Select(driver.find_element_by_xpath("//select[@name ='fromTo']"))
                select.select_by_value(str(i))
            sleep(1)
            btn = driver.find_element_by_xpath("//input[@name = 'btnQuery' and @value = '查詢']")
            btn.click()
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            df_data = df_data.append(self.clean_data(soup))
            sleep(5)
        ##-----------------------把不要的基金類型刪掉-----------------------------
        df_data = df_data.loc[df_data['基金類型'] != "指數股票型"]
        df_data = df_data.loc[df_data['基金類型'] != "貨幣市場基金"]
        ##-----------------------決定要找甚麼樣的前10大---------------------------
        df_data = df_data.sort_values(by=desire_type,ascending=False)
        df_data = df_data.reset_index(drop = True)
        df_result = df_data[df_data.index < number_of_funds]
        df_result = df_result.drop(columns=["序號"])
        self.df =df_result
        return self.df
    def clean_data(self,soup):   
        i = soup.select("table")[6].text.find("單位：新台幣百萬元")
        text = soup.select("table")[6].text[i+10:]
        i = text.find("1")
        text_list = text[i:].split("\n")
        df = pd.DataFrame(columns = ["序號", "基金名稱", "基金類型", "淨申贖金額", "申購金額", "買回金額", "基金規模", "最新淨值日期","淨值"])
        for i in range(0, len(text_list), 13):
            if len(text_list[i]) > 0 :
                df = df.append({"序號": text_list[i],
                                "基金名稱": text_list[i+1],
                                "基金類型": text_list[i+2],
                                "淨申贖金額": text_list[i+3],
                                "申購金額": text_list[i+4],
                                "買回金額": text_list[i+5],
                                "基金規模": text_list[i+6],
                                "最新淨值日期": text_list[i+7],
                                "淨值": text_list[i+8],
                }, ignore_index=True)
        return df
    def output_file(self):
        self.df.to_excel(r"/Users/chen-lichiang/Desktop/output_file/fund_crawler_output_file.xlsx",index=False)
        print(self.df)
#-----------------------------------use fund crawler class-----------------------------------------
crawler=funds_crawler(r"//usr/local/bin/chromedriver")
crawler.get_data(desire_type="淨申贖金額",desir_page=7,number_of_funds=10)
crawler.output_file()
