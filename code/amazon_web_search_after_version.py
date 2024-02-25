from bs4 import BeautifulSoup
from selenium import webdriver
import time
import urllib.request
import pandas as pd
from openpyxl.drawing.image import Image
from selenium.webdriver.common.by import By
import openpyxl
import math

select_num = input('''
    1.Amazon Devices & Accessories     2.Amazon Launchpad            3.Amazon Pantry
    4.Appliances                       5.Apps & Games                6.Arts, Crafts & Sewing       
    7.Audible Books & Originals        8.Automotive                  9.Baby                        
    10.Beauty & Personal Care          11.Books                      12.CDs & Vinyl                
    13.Camera & Photo                  14.Cell Phones & Accessories  15.Clothing, Shoes & Jewelry  
    16.Collectible Currencies          17.Computers & Accessories    18.Digital Music              
    19.Electronics                     20.Entertainment Collectibles 21.Gift Cards                 
    22.Grocery & Gourmet Food          23.Handmade Products          24.Health & Household         
    25.Home & Kitchen                  26.Industrial & Scientific    27.Kindle Store               
    28.Kitchen & Dining                29.Magazine Subscriptions     30.Movies & TV                
    31.Musical Instruments             32.Office Products            33.Patio, Lawn & Garden       
    34.Pet Supplies                    35.Software                   36.Sports & Outdoors               
    37.Sports Collectibles             38.Tools & Home Improvement   39.Toys & Games                    
    40.Video Games

    1.위 분야 중에서 자료를 수집할 분야의 번호를 선택하세요: ''')

width_num = int(input('    2.해당 분야에서 크롤링 할 건수는 몇건입니까?(1-100 건 사이 입력) : '))
save_data = input('    3.파일을 저장할 폴더명만 쓰시오.(예:c:\\data2\\) : ')
real_num = math.ceil(width_num / 50)

url = 'https://www.amazon.com/bestsellers?ld=NSGoogle'  # 주소
driver = webdriver.Edge()  # 엣지드라이버
driver.get(url=url)  # url 설정

driver.implicitly_wait(10)  # 대기

if select_num == '1':
    driver.find_element(By.XPATH,
                        "/html/body/div[1]/div[1]/div[2]/div/div/div/div[2]/div/div[2]/div/div/div[2]/div[1]").click()
else:
    str1 = f"/html/body/div[1]/div[1]/div[2]/div/div/div/div[2]/div/div[2]/div/div/div[2]/div[{select_num}]/a"
    driver.find_element(By.XPATH, str1).click()


def scroll_down(drivers):
    drivers.execute_script("window.scrollBy(0,9300);")
    time.sleep(1)


scroll_down(driver)

driver.implicitly_wait(3)

n = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (n.tm_year, n.tm_mon, n.tm_mday, n.tm_hour, n.tm_min, n.tm_sec)
ranking = []
name = []
price = []
review = []
point = []
src = []
img = []
count = 0
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

driver.implicitly_wait(3)

var1 = soup.find('div', class_='p13n-desktop-grid').find_all('div', class_='_cDEzb_iveVideoWrapper_JJ34T')

for i2 in var1:
    try:
        ranking.append(i2.find('span', class_='zg-bdg-text').get_text())
    except AttributeError:
        ranking.append('Nothing')
    try:
        name.append(i2.find('div', class_='_cDEzb_p13n-sc-css-line-clamp-3_g3dy1').get_text())
    except AttributeError:
        name.append('Nothing')
    try:
        price.append(i2.find('span', class_='p13n-sc-price').get_text())
    except AttributeError:
        price.append('Nothing')
    try:
        review.append(i2.find('span', class_='a-size-small').get_text())
    except AttributeError:
        review.append('Nothing')
    try:
        point.append(i2.find('span', class_='a-icon-alt').get_text())
    except AttributeError:
        point.append('Nothing')
    try:
        img.append(i2.find('img').attrs['src'])
    except AttributeError:
        img.append('Nothing')
    count += 1
    if count == width_num:
        break

if width_num > 50:
    driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div[2]/div[2]/ul/li[4]/a').click()
    scroll_down(driver)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    var1 = soup.find('div', class_='p13n-desktop-grid').find_all('div', class_='_cDEzb_iveVideoWrapper_JJ34T')
    for i2 in var1:
        try:
            ranking.append(i2.find('span', class_='zg-bdg-text').get_text())
        except AttributeError:
            ranking.append('Nothing')
        try:
            name.append(i2.find('div', class_='_cDEzb_p13n-sc-css-line-clamp-3_g3dy1').get_text())
        except AttributeError:
            name.append('Nothing')
        try:
            price.append(i2.find('span', class_='p13n-sc-price').get_text())
        except AttributeError:
            price.append('Nothing')
        try:
            review.append(i2.find('span', class_='a-size-small').get_text())
        except AttributeError:
            review.append('Nothing')
        try:
            point.append(i2.find('span', class_='a-icon-alt').get_text())
        except AttributeError:
            point.append('Nothing')
        try:
            img.append(i2.find('img').attrs['src'])
        except AttributeError:
            img.append('Nothing')
        count += 1
        if count == width_num:
            break

for i in range(0, len(name)):
    f = open(save_data + s + '-' + 'amazon' + '.txt', 'a', encoding='utf-8')
    f.write('1.판매순위:' + ranking[i] + '\n')
    f.write('2.제품소개:' + name[i] + '\n')
    f.write('3.가격:' + price[i] + '\n')
    f.write('4.상품평 수:' + review[i] + '\n')
    f.write('5.평점:' + point[i] + '\n')
    f.write('\n')
    f.close()

amazon1 = pd.DataFrame()
amazon1['판매순위'] = pd.Series(ranking)
amazon1['제품소개'] = pd.Series(name)
amazon1['가격'] = pd.Series(price)
amazon1['상품평 수'] = pd.Series(review)
amazon1['평점'] = pd.Series(point)
amazon1.to_excel('c:\\data\\' + s + '-' + 'amazon.xlsx', index=False)
num = 0
for i in range(0, len(img)):
    num += 1
    urllib.request.urlretrieve(img[i], "c:\\img2\\" + str(num) + ".jpg")

book = openpyxl.load_workbook('c:\\data\\' + s + '-' + 'amazon.xlsx')
count2 = 2
for i in range(1, width_num + 1):
    sheet = book.active
    sheet.column_dimensions['B'].width = 40
    img = Image('C:\\img2\\' + str(i) + '.jpg')
    sheet.add_image(img, 'B%s' % count2)
    sheet.row_dimensions[i + 1].height = 500
    count2 += 1

book.save('c:\\data2\\' + s + '-' + '아마존.xlsx')
