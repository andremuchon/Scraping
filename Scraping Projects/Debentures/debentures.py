import os
import winreg
import shutil
import time
from selenium import webdriver
from datetime import date, timedelta

# Dictionary to convert dates and get filenames
convert_dict = {
    '01': 'jan',
    '02': 'fev',
    '03': 'mar',
    '04': 'abr',
    '05': 'mai',
    '06': 'jun',
    '07': 'jul',
    '08': 'ago',
    '09': 'set',
    '10': 'out',
    '11': 'nov',
    '12': 'dez'
} 

# Function to convert date to standard filenames
def date_convert(ds):
    month = convert_dict[ds[2:4]]
    return 'd'+ds[-2:] + month + ds[:2] + '.xls'

# Function to find the downloads folder path
def get_download_path():
    sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
    return location

# Getting the dates
n = input("Quantos dias úteis vamos buscar (máximo de 5 dias úteis):")
print('\n')
days = []
for i in range(int(n)):
    l = input("{0}o dia útil (formato dd/mm/yyyy)".format(i+1))
    days.append(l)

# Adjusting the format
days = [d.replace('/','') for d in days]

# Opening the page on Chrome
url = 'https://www.anbima.com.br/informacoes/merc-sec-debentures/default.asp'
driver = webdriver.Chrome()
driver.get(url)

# Wait for the page to load
time.sleep(5)

# Loop through dates
for d in days:
    datainput = driver.find_element_by_xpath('//*[@id="cinza50"]/form/div/fieldset/table/tbody/tr/td/input[2]')
    datainput.clear()
    datainput.send_keys(d)
    try:
        consultbut = driver.find_element_by_xpath('//*[@id="cinza50"]/form/div/table/tbody/tr/td/img')
        consultbut.click()
        downbut = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/div/table[2]/tbody/tr/td/table/tbody/tr/td/a[1]')
        downbut.click()
    except:
        pass
    driver.execute_script('window.history.go(-1)')
    time.sleep(5)

# Close window
driver.quit()

# Locations
down_files = list(map(date_convert,days))
down = get_download_path()
server_path = 'DESTINATION PATH'

# Move files to destination folder
for f in down_files:
    file_path = os.path.join(down,f)
    server_dest_path = os.path.join(server_path,f)
    try:
        filemove = shutil.copyfile(file_path, server_dest_path)
        os.remove(file_path)
    except:
        print('Problema com o arquivo {0}'.format(f))
    
