from bs4 import BeautifulSoup
import requests
from selenium import webdriver

# url 가져오는 부
f = open("url.txt", "r")
lines = f.read().split('\n')

# 크롬 드라이버 가져오
driver = webdriver.Chrome('/Users/bonghayun/Desktop/project/python/python_excel_miele/chromedriver')

# url가져와서 돌리기
# lines에는 url.txt에서 가져온 url들이 배열형태로 담겨있음,
for n,index in enumerate(lines):
    url = index.replace('\n','')

    driver.get(url)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    # print(soup)
    cnt = len(soup.find_all('div', class_='basicList_title__3P9Q7'))

    for i in range(0,cnt):
        metadata = soup.find_all('div', class_='basicList_title__3P9Q7')[i]
        title = metadata.a.get('title')
        print("<제품명> : ", title)               # title

        price = soup.find_all('span', class_='price_num__2WUXn')[i].text
        print("<가격> : ", price)                # 가격

        url = metadata.a.get('href')
        print("<url> : ", url)                  # url

        print("===================================================")

driver.close()

