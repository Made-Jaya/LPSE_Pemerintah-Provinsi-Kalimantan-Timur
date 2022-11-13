from selenium import webdriver
import pandas as pd
import time
import os  
from tqdm import tqdm
from selenium.webdriver.common.by import By

chromedriver_location = "C:/Users/madej/Downloads/chromedriver_win32/chromedriver.exe"


driver = webdriver.Chrome(chromedriver_location)
driver.get('https://lpse.kaltimprov.go.id/eproc4/lelang')

halamanutama_all = pd.DataFrame()

driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
banyakhalaman = driver.find_element(By.XPATH, '//*[@id="tbllelang_paginate"]/ul/li[9]/a').text


def loaddata(i):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    driver.switch_to.window(driver.window_handles[0])
    link = '//*[@id="tbllelang"]/tbody/tr[{}]/td[2]/p[1]/a'.format(i)
    driver.find_element_by_xpath(str(link)).click()
    driver.switch_to.window(driver.window_handles[1])
    Calon_Pengumuman = pd.DataFrame(pd.read_html(driver.page_source)[0])
    driver.find_element_by_xpath(str('//*[@id="main"]/ul/li[2]/a')).click()
    Peserta = pd.DataFrame(pd.read_html(driver.page_source)[0])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
        
    Pengumuman = pd.DataFrame()
    Pengumuman['Deskripsi'] = Calon_Pengumuman[Calon_Pengumuman.columns[1]]
    Pengumuman.index = Calon_Pengumuman[Calon_Pengumuman.columns[0]]
    
    if len(Peserta.columns) > 2 :
        Peserta['Harga Penawaran'] = Peserta['Harga Penawaran'].str.replace('Rp. ', '')
        Peserta['Harga Penawaran'] = Peserta['Harga Penawaran'].str.replace('.', '')
        Peserta['Harga Penawaran'] = Peserta['Harga Penawaran'].str.replace(',', '.')
        Peserta['Harga Terkoreksi'] = Peserta['Harga Terkoreksi'].str.replace('Rp. ', '')
        Peserta['Harga Terkoreksi'] = Peserta['Harga Terkoreksi'].str.replace('.', '')
        Peserta['Harga Terkoreksi'] = Peserta['Harga Terkoreksi'].str.replace(',', '.')
    
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter("LPSE_POV_KALTIM/"+str(Pengumuman[Pengumuman.columns[0]][0])+'.xlsx', engine='xlsxwriter')
    Pengumuman.to_excel(writer, sheet_name='Pengumuman')
    Peserta.to_excel(writer, sheet_name='Peserta')
    writer.save()
    

    for i in tqdm(range(2,int(banyakhalaman))):
        time.sleep(5)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        halamanutama = pd.DataFrame(pd.read_html(driver.page_source)[1])
        halamanutama_all = halamanutama_all.append(halamanutama, ignore_index=True)
        driver.find_element_by_xpath(str('//*[@id="tbllelang_next"]/a')).click()
        for ii in range(1,len(pd.read_html(driver.page_source)[1])):
            try:
                loaddata(ii)
            except:
                print("Erorr"+str(ii))

writer = pd.ExcelWriter("LPSE_POV_KALTIM/halamanutama_all"+'.xlsx', engine='xlsxwriter')
halamanutama_all.to_excel(writer, sheet_name='Pengumuman')
halamanutama_all.to_excel(writer, sheet_name='Peserta')
writer.save()
    