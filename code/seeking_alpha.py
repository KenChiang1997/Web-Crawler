import time
import numpy as np 
import pandas as pd 
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


class crawler():
    """
    selenium_path 
    ticker -> string
    year -> string
    quarter -> string list 
    """
    def __init__(self,selenium_path,ticker,year,quarter):
        self.selenium_path = selenium_path 
        self.ticker = ticker 
        self.year = year 
        self.quarter = quarter 
        self.driver = webdriver.Chrome(self.selenium_path)
    def mian_fnc(self):
        self.driver.get("https://seekingalpha.com/")
        self.crawler_1()
        for q in self.quarter: 
            df = self.crawler_2( quarter=q, year=str(self.year) )
            self.driver.back()
            time.sleep(2)
            self.output_df(df,q=q,year=self.year)
        self.driver.quit()
    def crawler_1(self): 
        search = self.driver.find_element_by_xpath("//header/div[1]/div[1]/div[1]/div[1]/div[1]/input[1]")
        search.clear()
        search.send_keys( str(self.ticker) )
        search.send_keys(Keys.RETURN)
        time.sleep(2)
        _ = self.driver.find_element_by_xpath("//a[contains(text(),'Transcripts')]").click()
        time.sleep(2)
    def crawler_2(self,quarter,year):
        _ = self.driver.find_element_by_xpath("//a[contains(text(),'Apple Inc. (AAPL) CEO Tim Cook on "+str(quarter)+" "+str(year)+" Results ')]").click()
        time.sleep(5)
        soup = BeautifulSoup(self.driver.page_source, 'html.parser')
        df = self.clean_soup(soup)
        return df 
    def clean_soup(self,soup):
        participate = soup .find('div',{ 'role' : 'presentation' })
        df = pd.DataFrame()
        for i in participate.find_all('p') :
            series = pd.Series( i.text )
            df = df.append(series, ignore_index=True)
        return df
    def output_df(self,df,q,year):
        df.to_excel(r"/Users/chen-lichiang/Desktop/output_file/"+str(self.ticker)+"_"+str(year)+"_"+str(q)+".xlsx")  


seeking_alpha = crawler(selenium_path=r"//usr/local/bin/chromedriver",ticker="aapl",year="2020",quarter=["Q1","Q2","Q3","Q4"])
seeking_alpha.mian_fnc()