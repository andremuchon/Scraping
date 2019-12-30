import os
import numpy as np
import pandas as pd
import requests
import scipy.interpolate
import shutil
import winreg
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys


def get_download_path():
    """
    Returns the Downloads folder path
    """
    sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
    return location

def parsing(l):
    """
    Deals with the string format of the table
    """
    l_p = []
    for i in range(0,len(l),2):
        l_p.append(l[i:i+2])
    return l_p

def get_di_date(d):
    """
    Pulls the DIxIPCA data from table available at the URL
    """
    ipca_but = driver.find_element_by_name('slcTaxa')
    ipca_but.click()
    ipca_but.send_keys('DI x IPCA')
    ipca_but.send_keys(Keys.ENTER)
    data_area = driver.find_element_by_xpath('//*[@id="Data"]') # find date input area
    data_area.clear()
    data_area.send_keys(d) # cleans date field and inserts our date
    okbutton = driver.find_elements_by_xpath('//*[@id="divContainerIframeBmf"]/form/div/div/div[1]/div[2]/div/div[2]') # get the ok button
    okbutton[0].click() # click on the ok button
    w_url = requests.get(str(driver.current_url)).text 
    soup = BeautifulSoup(w_url, 'lxml') # use beautiful soup to deal with the html of the current url we loaded at our date
    infoslist = []
    for tr in soup.findAll("table"):
        for td in tr.find_all("td"):
            if not td.attrs.get('style'):
                infoslist.append(td.text) # get the elements of the table we need
    infoslist_p = parsing(infoslist) # use parsing function to reorganize
    di = pd.DataFrame(infoslist_p) # create the data frame with the information
    rename_cols = {0: 'Vertices', 1: 'DIxIPCA 252'} 
    di.rename(columns = rename_cols, inplace = True) # renaming the columns accordingly
    di = di.apply(lambda x: x.str.replace(',','.'))
    di['Vertices'] = pd.to_numeric(di['Vertices'], errors = 'coerce')
    di['DIxIPCA 252'] = pd.to_numeric(di['DIxIPCA 252'], errors = 'coerce')
    return di
    

def y(df,x):
    """
    Returns the cubic b-spline representation of the data in the Data Frame evaluated at x.
    """
    x_p=np.array(df['Vertices'])
    y_p=np.array(df['DIxIPCA 252'])
    cs = scipy.interpolate.splrep(x_p,y_p)
    return scipy.interpolate.splev(x,cs) 

def cubicspline(df):
    """
    Returns a Data Frame containing the evaluation of the cubic b-spline for all integers corresponding to days.
    """
    listDI = [ [n,y(df,n)] for n in range(1,max(df['Vertices']))]
    df_spline = pd.DataFrame(listDI, columns = ['Days', 'DIxIPCA 252'])
    return df_spline


wd = input("Digite a data da curva (dia útil!) no formato DD/MM/YYYY, para a curva mais recente, colocar o último dia útil\n")

url = 'http://www2.bmf.com.br/pages/portal/bmfbovespa/lumis/lum-taxas-referenciais-bmf-ptBR.asp'
driver = webdriver.Chrome()
driver.get(url) # OPENS URL


df_d = get_di_date(wd)
df_d = cubicspline(df_d) #MAIN TASK

file_path = os.getcwd() + '\\outDIxIPCA.xlsx'
df_d.to_excel(file_path, engine = 'openpyxl')
wd_adapt = wd.replace('/','-')
server_dest_path = get_download_path() + r'\DIxIPCA - Dia {0}.xlsx'.format(wd_adapt)
filemove = shutil.copyfile(file_path, server_dest_path)

driver.quit()
print('Done!')
