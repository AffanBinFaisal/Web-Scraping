from bs4 import BeautifulSoup
from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from openpyxl import Workbook
import time
import openpyxl
filename="smartphones.txt"
wb = openpyxl.Workbook()
sheet = wb.active   
c = sheet['A1'] 
c.value = "Title"
c = sheet['B1'] 
c.value = "Price"
c = sheet['C1'] 
c.value = "Discounted Price"
c = sheet['D1'] 
c.value = "Discount%"
c = sheet['E1']
c.value = "Deivery Details"
c = sheet['F1'] 
c.value = "Deivery Time"
c = sheet['G1'] 
c.value = "Deivery Fee"
c = sheet['H1'] 
c.value = "Main Image Link"
c = sheet['I1'] 
c.value = "Brand"
c = sheet['J1'] 
c.value = "Score"
c = sheet['K1'] 
c.value = "Max Score"
c = sheet['L1'] 
c.value = "Reviews Count"
c = sheet['M1'] 
c.value = "Reviews"
c = sheet['N1'] 
c.value = "Vendor"
c = sheet['O1'] 
c.value = "Specifications"
c = sheet['P1'] 
c.value = "Slug" 
c = sheet['Q1'] 
c.value = "Category"   

i = 1
file = open(filename, "r")
s = Service('./chromedriver.exe')
driver = webdriver.Chrome(service=s)
driver.maximize_window()
url = file.readline()

while url:
    i += 1
    driver.get(url)
    driver.implicitly_wait(15)
    page_src = driver.page_source

    soup = BeautifulSoup(page_src, 'html.parser')

    try:
        title = soup.find('span', class_ = 'pdp-mod-product-badge-title').text
        c = sheet['A'+str(i)] 
        c.value = title
    except:
        pass
    
    try:
        oprice = soup.find('span', class_ = 'pdp-price pdp-price_type_deleted pdp-price_color_lightgray pdp-price_size_xs').text
        c = sheet['B'+str(i)] 
        c.value = oprice
    except:
        pass
    
    try:
        dprice = soup.find('span', class_ = 'pdp-price pdp-price_type_normal pdp-price_color_orange pdp-price_size_xl').text
        c = sheet['C'+str(i)] 
        c.value = dprice
    except:
        pass
    
    try:
        discount = soup.find('span', class_ = 'pdp-product-price__discount').text
        c = sheet['D'+str(i)] 
        c.value = discount[1:]
    except:
        pass
    
    try:
        dedetails_list = []
        dedetails = soup.findAll('div', class_ = 'delivery-option-item__title')
        for dedetail in dedetails:
            dedetails_list.append(dedetail.text)
        c = sheet['E'+str(i)] 
        c.value = str(dedetails_list)
    except:
        pass
    
    try:
        deltime = soup.find('div', class_ = 'delivery-option-item__time').text
        c = sheet['F'+str(i)] 
        c.value = deltime
    except:
        pass
    
    try:
        delfee = soup.find('div', class_ = 'delivery-option-item__shipping-fee').text
        c = sheet['G'+str(i)] 
        c.value = delfee
    except:
        pass
    
    try:
        img = soup.find('img', class_ = 'pdp-mod-common-image gallery-preview-panel__image').get("src")
        c = sheet['H'+str(i)] 
        c.value = img
    except:
        pass
    
    try:
        brand = soup.find('a', class_ = 'pdp-link pdp-link_size_s pdp-link_theme_blue pdp-product-brand__brand-link').text
        c = sheet['I'+str(i)] 
        c.value = brand
    except:
        pass
    
    try:
        scoreavg = soup.find('span', class_ = 'score-average').text
        c = sheet['J'+str(i)] 
        c.value = scoreavg
    except:
        pass
    
    try:
        scoremax = soup.find('span', class_ = 'score-max').text
        c = sheet['K'+str(i)] 
        c.value = scoremax[1:]
    except:
        pass
    
    try:
        count = soup.find('div', class_ = 'count').text
        c = sheet['L'+str(i)] 
        c.value = count[:-8]
    except:
        pass
    
    try:
        reviews_list = []
        reviews = soup.findAll('div', class_ = 'content')[1:]
        for review in reviews:
            reviews_list.append(review.text)
        c = sheet['M'+str(i)] 
        c.value = str(reviews_list)
    except:
        pass
    
    c = sheet['N'+str(i)] 
    c.value = "https://www.daraz.pk/#"
    
    try:
        details = soup.find('div', class_ = 'html-content pdp-product-highlights').text
        c = sheet['O'+str(i)] 
        c.value = details
    except:
        pass
    
    c = sheet['P'+str(i)] 
    c.value = url
    
    c = sheet['Q'+str(i)] 
    c.value = "TV"
    
    url = file.readline()
    
wb.save("h.xlsx")
