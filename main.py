from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import xlsxwriter
from time import sleep
import re

# 크롬 드라이버 가져오기
driver = webdriver.Chrome('./chromedriver')

# 엑셀 파일 지정 & 쉬트지 열기
workbook = xlsxwriter.Workbook('price.xlsx')
worksheet = workbook.add_worksheet()

#나중에 쓸 셀 형식들 (색상, 보더, 글 색상 등 )
data_format_first_line = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 2, 'bold': True})
data_format1 = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 1})

#수집한 정보를 입력
row = 1

#첫 줄 내용
row_one_line = ["MODEL", "SP", "Promotion Price", "Differ",
                "가격_1", "쇼핑물_1", "url_1", "상품명_1",
                "가격_2", "쇼핑물_2", "url_2", "상품명_2",
                "가격_3", "쇼핑물_3", "url_3", "상품명_3",
                "가격_4", "쇼핑물_4", "url_4", "상품명_4",
                "가격_5", "쇼핑물_5", "url_5", "상품명_5",
                "가격_6", "쇼핑물_6", "url_6", "상품명_6",
                "가격_7", "쇼핑물_7", "url_7", "상품명_7",
                "가격_8", "쇼핑물_8", "url_8", "상품명_8",
                "가격_9", "쇼핑물_9", "url_9", "상품명_9",
                "가격_10", "쇼핑물_10", "url_10", "상품명_10"]

#첫 줄 형식지정 & 틀 고정
worksheet.set_row(0, cell_format=data_format_first_line)
worksheet.freeze_panes(1, 1)

#정보가 들어갈 2차원 배열, url.txt에서 가져온 정보를 여기에 주입할 것임
data = []

# url.txt 파일이을 열고 띄어쓰기 단위로 모델명 가격1 가격2 url을 가져옴
f = open("url_test.txt", "r")
lines = f.read().split('\n')
f.close()

# url.txt에서 가져온 정보가 [제품명 가격1 가격2 url, 제품명 가격1 가격2 url, ...] 이런식으로 저장되어있음,
# index에는 "제품명 가격1 가격2 url" 이렇게 들어있고
for n,index in enumerate(lines):
    #배송비를 가져오기 위한 0, 왜냐면 쇼핑물 모음집이 있으면 배열이 댕겨져서, 안맞아.
    #그걸 보정하기 위해서 쇼핑물 모음집이 있으면 그 수 만큼 댕겨줄꺼
    mall_group = 0

    # lines[0] 의 정보 중에서 불필요한 \n을 제거
    info_product = index.replace('\n','')

    #2차원 데이터를 위해 선언한 data에 [[제품명, 가격1, 가격2, url], ...] 가 되도록 계속해서 추가함
    data.insert(len(data), [info_product.split(" ")[0], info_product.split(" ")[1], info_product.split(" ")[2], info_product.split(" ")[3]])

    #저장된 data에서 가장 최신 정보 중 url에 접속해서 스크롤을 아래로 쭉 해(html파일 로딩을 위해서)
    driver.get(data[n][3])
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    #컴퓨터가 html파일을 전부 가져오지 못해서 html을 전부 가져오기 위해 0.1초의 시간을 줌!
    sleep(0.5)
    html = driver.page_source

    soup = BeautifulSoup(html, 'html.parser')

    #크롤링한 html파일에 가격정보가 몇개가 들어있는지 파악
    price_info_numbers = len(soup.find_all('div', class_='basicList_title__3P9Q7'))

    #데이터의 갯수를 제한하는 단계
    if(price_info_numbers > 10):
        price_info_numbers = 10
    else:
        pass

    #가독성을 위해서 정보가 총 차지하는 셀에 격으로 색을 칠함
    for i in range(30):
        index = (4 * i - 3) + 3
        if (i % 2) != 0:
            worksheet.set_column(index, index + 3, cell_format=data_format1)

    # print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>___" + data[n][0] + "___<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<" )

    #크롤링한 정보를 담을 data_arr이라는 배열,
    #한 정보의 싸이클이 끝나고 다시 시작하면 초기화.
    data_arr = []

    #price_info_numbers는 갯수일 뿐.
    for i in range(0,price_info_numbers):

        metadata = soup.find_all('div', class_='basicList_title__3P9Q7')[i]
        title = metadata.a.get('title')
        # print("<제품명> : ", title)  # title

        price = soup.find_all('span', class_='price_num__2WUXn')[i].text.replace(',', '').replace('원', '')
        # print("<가격> : ", price)  # 가격

        url = metadata.a.get('href')
        # print("<url> : ", url)  # url

#원래 배송비 가져오는 위치



        #위에서 크롤링한 html에서 정보 뺴오자
        if(soup.find_all('a', class_='basicList_mall__sbVax')[i].img):
            #일반적인 쇼핑물
            mall_name = soup.find_all('a', class_='basicList_mall__sbVax')[i].img.get('alt')
            # print("<쇼핑물> : ", mall_name)                # 쇼핑물
            #위에서 가져온 데이터들을 담아버려~
            #이 배열도 2차원임. 한 제품의 정보들 중에 한 쇼핑물에 가격등 정보를 [0][n]에 담는 것임
            data_arr.insert(len(data_arr),[int(price),mall_name,url,title])

        else:
            # 네이버 스토어나 가격비교 원부
            mall_name = soup.find_all('a', class_='basicList_mall__sbVax')[i].text
            # print("<쇼핑물> : ", mall_name)                # 쇼핑물

            # data_arr.insert(len(data_arr), [int(price), "err" + mall_name, url, title])

            # 해당 쇼핑물에 배송비가 있는지
            # 배송비 > 배송정보 숫자만 받기 > 숫자 변수에 저장,
            shipping_price = False
            if(mall_name != '쇼핑몰별 최저가'):
                shipping_price_box = soup.find_all('ul', class_='basicList_mall_option__1qEUo')[i - mall_group]
                shipping_price = shipping_price_box.find_all('em', class_='basicList_option__3eF2s')[0].text.replace(',','')

                if (re.search(r'\d+', shipping_price)):
                    shipping_price = int(re.search(r'\d+', shipping_price).group())
                    # print(shipping_price)
                    pass
                else:
                    shipping_price = False
            else:
                #쇼핑물별 최저가인 경우
                #n번째를 추가해서 배열 순서를 맞추자
                mall_group = mall_group + 1


            if shipping_price:
                # 네이버 스토어 중에서 배송비가 있는 곳
                # print(data[n][0]+str(i) + mall_name + ' > ' + str(shipping_price))
                data_arr.insert(len(data_arr),[int(price) + shipping_price, "네이버["+mall_name + "-배송비 포함]", url, title])

            else:
                # 가격비교 원부 or 배송비 없는 네이버 스토어
                if mall_name == '쇼핑몰별 최저가':
                    pass
                else:
                    data_arr.insert(len(data_arr),[int(price),'네이버[' + mall_name + ']',url,title])



    #한 제품의 모든 쇼핑물에 대한 정보를 전부 담고 있는 data_arr을 출력
    data_arr.sort(key=lambda x: x[0])
    print(data_arr)

    # 데이터를 입력할 첫번째 컬럼
    # 모든 제품의 데이터 정보는 n번째 줄, 5번째 칼럼에서 시작해야함
    # 그래서 엑셀에 기록하기전에 항상 4로 초기화
    col = 4

    # data_arr에는 한 제품에 대한 모든 쇼핑물의 정보가 들어있으니깐, 쇼핑물 하나하나 엑셀에다가 정리하자
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


#3번쨰 컬럼 서식 (differ 값 구하는 식)
for row_num in range(1, 100):
    worksheet.write_formula(row_num-1, 3,
                            '= -$C%d + $E%d' % (row_num, row_num))

#differ값이 0보다 작으면 아래 형식 지정
differ_smaller_than_zero = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

#첫 줄에 데이터 입력
for n, info in enumerate(row_one_line):
    worksheet.write(0, n, info)

worksheet.conditional_format('D2:D61', {'type': 'cell',
                                         'criteria': '<',
                                         'value': 0,
                                         'format': differ_smaller_than_zero})


workbook.close()
driver.close()

