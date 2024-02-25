# Amazon_Web_Search_Project

## 프로젝트 수행기간 : 2020.04.10~2020.05.28

SNS빅데이터 분석 개인프로젝트 - 아마존 웹 사이트 검색 프로그램

![image](https://github.com/shinho123/Amazon_Web_Search_Project/assets/105840783/161d6a16-bdec-4ee8-9c73-f63039b1f628)

### 수행 내용 

1. 아마존 웹사이트(https://www.amazon.com/bestsellers?ld=NSGoogle)에서 원하는 카테고리 번호 선택(총 40개)

```python
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
```
2. 크롤링 할 건수 입력(최소 1건 ~ 최대 100건)
```python
width_num = int(input('    2.해당 분야 에서 크롤링 할 건수는 몇건 입니까?(1-100 건 사이 입력) : '))
```


3. 파일을 저장할 폴더명 입력(저장된 시간을 함께 표시)

```python
save_data = input('    3.파일을 저장할 폴더명 : ')
```

### 텍스트 및 이미지 데이터는 data, img2에 각각 저장되며 각각 저장된 데이터가 병합되어 data2에 저장됨

#### data

```python
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
```

![image](https://github.com/shinho123/Amazon_Web_Search_Project/assets/105840783/3d8a2776-2c51-436e-b3cd-f101f8e12757)

#### img2

```python
num = 0
for i in range(0, len(img)):
    num += 1
    urllib.request.urlretrieve(img[i], "c:\\img2\\" + str(num) + ".jpg")
```

![image](https://github.com/shinho123/Amazon_Web_Search_Project/assets/105840783/7291f23d-7246-49f0-a9f8-0557086ff0bb)

#### data2

```python
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
```

![image](https://github.com/shinho123/Amazon_Web_Search_Project/assets/105840783/14831c4c-08b7-4188-957c-35ca0a9ea1e9)




