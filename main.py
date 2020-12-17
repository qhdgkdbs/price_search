from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import xlsxwriter

# url 가져오는 부
f = open("url.txt", "r")
lines = f.read().split('\n')

# 크롬 드라이버 가져오
driver = webdriver.Chrome('/Users/bonghayun/Desktop/project/python/python_excel_miele/chromedriver')

workbook = xlsxwriter.Workbook('price.xlsx')
worksheet = workbook.add_worksheet()
row = 1

data_format_first_line = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 2, 'bold': True})
data_format1 = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 1})

row_one_line = ["MODEL", "SP", "Promotion Price", "Differ",
                "가격_1", "쇼핑물_1", "url_1", "상품명_1",
                "가격_2", "쇼핑물_2", "url_2", "상품명_2",
                "가격_3", "쇼핑물_3", "url_3", "상품명_3",
                "가격_4", "쇼핑물_4", "url_4", "상품명_4",
                "가격_5", "쇼핑물_5", "url_5", "상품명_5",
                "가격_6", "쇼핑물_6", "url_6", "상품명_6",
                "가격_7", "쇼핑물_7", "url_7", "상품명_7",
                "가격_8", "쇼핑물_8", "url_8", "상품명_8",
                "가격_9", "쇼핑물_9", "url_9", "상품명_9"]



worksheet.set_row(0, cell_format=data_format_first_line)
worksheet.freeze_panes(1, 1)

#3번쨰 컬럼 서식
for row_num in range(1, 61):
    worksheet.write_formula(row_num-1, 3,
                            '= -$C%d + $E%d' % (row_num, row_num))

    # Write some data for the formula.

#첫 줄에 데이터 입력
for n, info in enumerate(row_one_line):
    worksheet.write(0, n, info)

data = []

# url가져와서 돌리기
# lines에는 url.txt에서 가져온 url들이 배열형태로 담겨있음,
for n,index in enumerate(lines):

    info_name_url = index.replace('\n','')

    data.insert(len(data), [info_name_url.split(" ")[0], info_name_url.split(" ")[1], info_name_url.split(" ")[2], info_name_url.split(" ")[3]])

    driver.get(data[n][3])
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    # print(soup)
    cnt = len(soup.find_all('div', class_='basicList_title__3P9Q7'))


    #데이터의 갯수를 10개로 제한하기
    if(cnt > 8):
        cnt = 8
    else:
        pass

    for i in range(30):
        index = (4 * i - 3) + 3
        if (i % 2) != 0:
            worksheet.set_column(index, index + 3, cell_format=data_format1)

    print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>___" + data[n][0] + "___<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<" )
    # 이게 그러니깐 한 품종이야, 그러면 데이터를 한번에 저장하려면, 위에다가 저장해야해



    data_arr = []

    for i in range(0,cnt):

        if(soup.find_all('a', class_='basicList_mall__sbVax')[i].img):
            mall_name = soup.find_all('a', class_='basicList_mall__sbVax')[i].img.get('alt')
            # print("<쇼핑물> : ", mall_name)                # 쇼핑물

            metadata = soup.find_all('div', class_='basicList_title__3P9Q7')[i]
            title = metadata.a.get('title')
            # print("<제품명> : ", title)  # title

            price = soup.find_all('span', class_='price_num__2WUXn')[i].text.replace(',','').replace('원','')
            # print("<가격> : ", price)  # 가격

            url = metadata.a.get('href')
            # print("<url> : ", url)  # url원

            # print("===================================================")
            data_arr.insert(len(data_arr),[price,mall_name,url,title])

        else:
            # print("종합가격 비교")
            # print("===================================================")
            pass
            # data_arr.insert(len(data_arr),[price,"가격비교가 최저가",url,"가격비교 최저"])

    print(data_arr)

    col = 4



    for info in data_arr:
        # worksheet.set_column(col + 1, cell_format=data_format2)
        worksheet.write(row, 0, data[n][0].upper())
        worksheet.write(row, 1, data[n][1])
        worksheet.write(row, 2, data[n][2])
        worksheet.write_row(row, col, info)
        col = col + 4

    #마지막에 셀에 스페이스 추가 for 깔끔
    worksheet.write_row(row, col, " ")

    row = row + 1

format1 = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

worksheet.conditional_format('D2:D61', {'type': 'cell',
                                         'criteria': '<',
                                         'value': 0,
                                         'format': format1})


workbook.close()
driver.close()

