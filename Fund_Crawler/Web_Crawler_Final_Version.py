# basic
import os
import re
import time
import warnings 
import numpy as np
from tqdm import tqdm 
import pandas as pd 
warnings.filterwarnings("ignore")

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


# excel
from module import Excel_Work,Address_DF


"""

(1.) If Google Chrome Update Version , u have to update selenium version too .  

    Selenium Chrome Driver Download Path : https://sites.google.com/chromium.org/driver/ 

(2.)  Under if __name__ == '__main__' :
    
    # ------------- Variable ------------- 
    Selenium     = "D:\My Documents\kenc\桌面\ken_python\ins客戶歸戶\chromedriver.exe"  --> Local Selenium Path Setting
    Type         = "offshore" # onshore --> lower cap                                  --> 境內基金觀測站 or 境外基金觀測站
    月份         =  "October"                                                           --> 月份 Selection ( Previous Month )
    Layer_2      = True # ( 點進去找基金各類級別資料 )                                    --> 點進去找法人級別 
    desire_page  = 10                                                                   --> Default set 10 , 總共爬 10 頁 (全部)                     
    # --------------------------------------- 
    

    # Call Function and Start Crawler 
    # 1.) offshore with 法人級別 
    Crawler                     = funds_crawler(selenium_path=Selenium,Type=Type,月份=月份,Layer_2=Layer_2)
    institution_data,df_data    = Crawler.get_data(desire_type="申購金額",desir_page=desire_page) # offhsore
    Crawler.To_Excel(df_data)

"""


#-----------------------------------function build------------------------------------------

class funds_crawler():

    def __init__(self,selenium_path,excel_output_path,Type,月份,Layer_2):

        self.selenium_path     = selenium_path
        self.excel_output_path = excel_output_path
        self.df            = pd.DataFrame()
        self.法人級別_df    = pd.DataFrame()
        
        self.月份     = 月份
        self.Type     = Type 
        self.Layer_2  = Layer_2

        if Type == "onshore" :
            self.url = "https://announce.fundclear.com.tw/MOPSonshoreFundWeb/INQ713.jsp"

        elif Type == "offshore":
            self.url = "https://announce.fundclear.com.tw/MOPSFundWeb/INQ712.jsp?commit=1&fundClassType=1&statType=1&fundAsset=all&fundInv=all&fundCurr=all&fromTo=1&dataRange=1"

    def get_data(self,desire_type,desir_page):

        # chrome driver setting
        options = webdriver.ChromeOptions()  
        options.add_argument("--incognito")
        options.add_argument("--start-maximized") 

        driver = webdriver.Chrome(self.selenium_path,options=options)
        driver.get(self.url)
        
        # select time range
        time_range =  Select(driver.find_element_by_xpath("//select[@name ='dataRange']"))
        time_range.select_by_value(str(1))

        #  start fetch data
        if desire_type == "淨申贖金額" :
            driver.find_element_by_xpath("//input[@name ='statType' and @value=3]").click() # 點淨申購資訊
        elif  desire_type == "申購金額" :
            driver.find_element_by_xpath("//input[@name ='statType' and @value=1]").click() # 點淨申購資訊

        df_data    = self.df
        法人級別_df = self.法人級別_df  
    
        for i in range(1,desir_page+1): #抓多少頁
            #顯示區間
            if i >= 2 :
                driver.find_element_by_xpath("//select[@name ='fromTo']").click()
                select = Select(driver.find_element_by_xpath("//select[@name ='fromTo']"))
                select.select_by_value(str(i)) 

            time.sleep(2)

            # 選完 PAGE 之後 Click 查尋 
            btn = driver.find_element_by_xpath("//input[@name = 'btnQuery' and @value = '查詢']")
            btn.click()

            # Locate Fund_links with xpath
            fund_links = driver.find_elements_by_xpath("//a[@href ='#']")
            fund_links = [fund_links[i] for i in range(0,len(fund_links),9) ]

        
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            if self.Type == "offshore" :

                data = self.offshore_clean_data(soup)

                if self.Layer_2 :
                    
                    # 點每一個基金 append 法人級的 data
                    for i in tqdm(range(data.shape[0])):

                        fund_name = data['基金名稱'][i]
                        fund_links[i].click()

                        time.sleep(2)
                        window_after  = driver.window_handles[1]
                        window_before = driver.window_handles[0]
                        driver.switch_to_window(window_after)

                        href       = driver.page_source
                        soup       = BeautifulSoup(href, 'html.parser')
                        法人級別_df = 法人級別_df.append( self.Clean_法人級別_Data(soup) ,ignore_index=True)

                        driver.close()
                        driver.switch_to_window(window_before)
                        time.sleep(2)

            elif self.Type == "onshore" :
                data = self.onshore_clean_data(soup)

            df_data = df_data.append(data)
            time.sleep(5)

        if self.Type == "offshore" :
            
            if self.Layer_2 :
                self.法人級別_df = 法人級別_df
                return  法人級別_df , df_data

            else :
                return df_data

        elif self.Type == "onshore" :

            return df_data

    def Clean_法人級別_Data(self,soup):   
        #　跳出視窗那個
        i    = soup.select("table")[0].text.find("單位：新台幣百萬元")
        text = soup.select("table")[0].text[i+10:]
        i    = text.find("1")
        text_list = text[i:].split("\n")
        text_list =  self.Address_Fund_Name(text_list)
        df = pd.DataFrame(columns = ["序號", "基金級別名稱", "基金類型", "申購金額", "買回金額", "淨申購金額","國人持有基金", "基金規模","最新淨值日期","淨值"])

        for i in range(0, len(text_list), 12):

            try:
                if len(text_list[i]) > 0 :
                    df = df.append({"序號"          : text_list[i],
                                    "基金級別名稱"   : text_list[i+1],
                                    "基金類型"       : text_list[i+2],   
                                    "申購金額"       : text_list[i+3],
                                    "買回金額"       : text_list[i+4],
                                    "淨申購金額"     : text_list[i+5],
                                    "國人持有基金"   : text_list[i+6],
                                    "基金規模"       : text_list[i+7],
                                    "最新淨值日期"   : text_list[i+8],
                                    "淨值"          : text_list[i+9],
                    }, ignore_index=True)
            except:
                pass 

        print(df)
        return df

    def Address_Fund_Name(self,list):

        Address_list    = [ list[i].replace("\t","") for i in range(len(list)) if list[i].replace("\t","") != ""]
        drop_index_list = []

        for i,element in enumerate(Address_list) :

            try:
                if "本基金" in Address_list[i+1] or "(基金之" in Address_list[i+1]  or "(基金主要" in Address_list[i+1] : 
                    Address_list[i] += Address_list[i+1]
                    drop_index_list.append(i+1)

            except:
                pass 
        
        Address_list = pd.DataFrame(Address_list,columns=['Text']).drop(drop_index_list)['Text'].to_list()
        return Address_list

    def offshore_clean_data(self,soup):   
        # offshore fund clean
        i         = soup.select("table")[6].text.find("單位：新台幣百萬元")
        text      = soup.select("table")[6].text[i+10:]
        i         = text.find("1")
        text_list = text[i:].split("\n")
        text_list =  self.Address_Fund_Name(text_list)

        df        = pd.DataFrame(columns = ["序號", "基金名稱", "基金類型", "淨申贖金額", "申購金額", "買回金額","國人持有基金", "基金規模","國人持有基金佔比","法定上限"])
        for i in range(0, len(text_list), 12):
            
            if len(text_list[i]) > 0 :
                df = df.append({"序號"             : text_list[i],
                                "基金名稱"          : text_list[i+1],
                                "基金類型"          : text_list[i+2],   
                                "申購金額"          : text_list[i+3],
                                "買回金額"          : text_list[i+4],
                                "淨申贖金額"        : text_list[i+5],
                                "國人持有基金"      : text_list[i+6],
                                "基金規模"         : text_list[i+7],
                                "國人持有基金佔比"  : text_list[i+8],
                                "法定上限"         : text_list[i+9],

                }, ignore_index=True)
        return df
    
    def onshore_clean_data(self,soup) :   

        # onshore fund clean
        i    = soup.select("table")[6].text.find("單位：新台幣百萬元")
        text = soup.select("table")[6].text[i+10:]
        i    = text.find("1")
        text_list = text[i:].split("\n")
        df = pd.DataFrame(columns = ["序號", "基金名稱", "基金類型", "淨申贖金額", "申購金額", "買回金額", "基金規模", "最新淨值日期","淨值"])

        for i in range(0, len(text_list), 13):


            if len(text_list[i]) > 0 :
                df = df.append({"序號"           : text_list[i],
                                "基金名稱"       : text_list[i+1],
                                "基金類型"       : text_list[i+2],
                                "淨申贖金額"     : text_list[i+3],
                                "申購金額"       : text_list[i+4],
                                "買回金額"       : text_list[i+5],
                                "基金規模"       : text_list[i+6],
                                "最新淨值日期"    : text_list[i+7],
                                "淨值"           : text_list[i+8],
                }, ignore_index=True)

        return df

    def To_Excel(self,df_data):

        Address_df = Address_DF()

        if self.Layer_2 : 
            
            法人級別_df     = self.法人級別_df
            法人級別_output = Address_df.法人級別種類(法人級別_df)
            

        申購金額_result ,買回金額_result ,淨申贖金額_result = Address_df.Top_20_DF(df_data)
        
        excel = Excel_Work()
        wb=excel.write_excel(df=df_data,line=True,tabel_style=False,tabel_style_name="TableStyleMedium9",excel_sheet_name='Fund Raw Data')
        wb.save(self.excel_output_path+"\\"+str(self.月份)+"_"+str(self.Type)+"_output.xlsx")
        excel.append_excel(excel_path=self.excel_output_path+"\\"+str(self.月份)+"_"+str(self.Type)+"_output.xlsx" ,df=申購金額_result   ,excel_sheet_name="總申購 TOP 20 raw data")
        excel.append_excel(excel_path=self.excel_output_path+"\\"+str(self.月份)+"_"+str(self.Type)+"_output.xlsx" ,df=買回金額_result   ,excel_sheet_name="總買回 TOP 20 raw data")
        excel.append_excel(excel_path=self.excel_output_path+"\\"+str(self.月份)+"_"+str(self.Type)+"_output.xlsx" ,df=淨申贖金額_result ,excel_sheet_name="淨申購 TOP 20 raw data")


        if self.Layer_2 :

            excel.append_excel(excel_path=self.excel_output_path+"\\"+str(self.月份)+"_"+str(self.Type)+"_output.xlsx" ,df=法人級別_df       ,excel_sheet_name="Account Raw Data")
            excel.append_excel(excel_path=self.excel_output_path+"\\"+str(self.月份)+"_"+str(self.Type)+"_output.xlsx" ,df=法人級別_output   ,excel_sheet_name="法人級別種類")


        print(" Task Complete ! ")

if __name__ == '__main__':

    # ------------- Variable ------------- 
    excel_output_path  = r"D:\My Documents\kenc\桌面\Ken_Final_Python\Ins客戶歸戶\基金觀測站\output"
    Selenium           = "D:\My Documents\kenc\桌面\ken_python\ins客戶歸戶\chromedriver.exe"
    Type               = "offshore" # onshore --> lower cap 
    月份               = "October"
    Layer_2            = True # ( 點進去找基金各類級別資料 )
    desire_page        = 1 
    # --------------------------------------- 

    # 1.) offshore with 法人級別 
    Crawler                     = funds_crawler(selenium_path=Selenium,excel_output_path=excel_output_path,Type=Type,月份=月份,Layer_2=Layer_2)
    institution_data,df_data    = Crawler.get_data(desire_type="申購金額",desir_page=desire_page) # offhsore
    Crawler.To_Excel(df_data)

    # 2.) offshore with 法人級別  with Layer_2 is None
    # Crawler    = funds_crawler(selenium_path=Selenium,excel_output_path=excel_output_path,Type=Type,月份=月份,Layer_2=Layer_2)
    # df_data    = Crawler.get_data(desire_type="申購金額",desir_page=desire_page) # offhsore
    # Crawler.To_Excel(df_data)


    # 3.) onshore
    # Crawler         = funds_crawler(selenium_path=Selenium,excel_output_path=excel_output_path,Type='onshore',月份=月份,Layer_2=None)
    # df_data         = Crawler.get_data(desire_type="申購金額",desir_page=2) # offhsore
    # Crawler.To_Excel(df_data)


