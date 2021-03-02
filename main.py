# -*- Encoding: UTF-8 -*- #

from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import xlsxwriter
from time import sleep
import re
import sys

def get_review(START_DATE, END_DATE, self):
    cannot_get = []

    # 크롬 드라이버 가져오기
    driver = webdriver.Chrome('./src/chromedriver')

    # 엑셀 파일 지정 & 쉬트지 열기
    # 제목에 날짜
    workbook_title = "./review/review_from_" +str(START_DATE) +"_to_"+ str(END_DATE) + ".xlsx"
    workbook = xlsxwriter.Workbook(workbook_title)
    worksheet = workbook.add_worksheet()

    #나중에 쓸 셀 형식들 (색상, 보더, 글 색상 등 )
    # data_format_first_line = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 2, 'bold': True})
    # data_format1 = workbook.add_format({'bg_color': '#E8F8FF', 'border': 7, 'bottom': 1})


    #수집한 정보를 입력
    row = 1

    #첫 줄 내용
    row_one_line = ["쇼핑물", "상품명", "작성자", "별점 수", "상품평", "사진URL","작성날짜"]

    #differ값이 0보다 작으면 아래 형식 지정
    rate_down_three = workbook.add_format({'bg_color': '#FFC7CE'})

    rate_up_three = workbook.add_format({'bg_color': '#A2D4F2'})

    title_cell = workbook.add_format({'bg_color': '#A2D4F2'})                               

    #첫 줄에 데이터 입력
    for n, info in enumerate(row_one_line):
        worksheet.write(0, n, info)

    worksheet.conditional_format('D2:D1000', {'type': 'cell',
                                            'criteria': '<=',
                                            'value': 3,
                                            'format': rate_down_three})
                                            
    worksheet.conditional_format('D2:D1000', {'type': 'cell',
                                            'criteria': '>',
                                            'value': 3,
                                            'format': rate_up_three})

    # worksheet.conditional_format('D3:D1000', {'type':'blanks', 'format': title_cell})


    #첫 줄 형식지정 & 틀 고정
    # worksheet.set_row(0, cell_format=data_format_first_line)
    worksheet.freeze_panes(1, 1)
    worksheet.set_column(4, 4, 60)

    #정보가 들어갈 2차원 배열, url.txt에서 가져온 정보를 여기에 주입할 것임
    data = []

    # url.txt 파일이을 열고 띄어쓰기 단위로 모델명 가격1 가격2 url을 가져옴
    # f = open("url_test.txt", "r")
    # lines = f.read().split('\n')
    # f.close()
    with open('./src/pd_info_test.txt', encoding='utf-8') as f:
        lines = f.read().split('\n')


    # url.txt에서 가져온 정보가 [제품명 원부코드, ...] 이런식으로 저장되어있음,
    # index에는 "제품명 가격1 가격2 url" 이렇게 들어있고
    for n,index in enumerate(lines):

        # lines[0] 의 정보 중에서 불필요한 \n을 제거
        info_product = index.replace('\n','')

        #2차원 데이터를 위해 선언한 data에 [[제품명 원부코드 원부url], ...] 가 되도록 계속해서 추가함
        data.insert(len(data), [info_product.split(" ")[0], info_product.split(" ")[1], "https://msearch.shopping.naver.com/catalog/" + info_product.split(" ")[1] + "/reviews"])
        product_review_page = data[n][2]
        
        #저장된 data에서 가장 최신 정보 중 url에 접속해서 스크롤을 아래로 쭉 해(html파일 로딩을 위해서)
        driver.get(product_review_page)
        sleep(0.3)

        
        try:
            driver.find_element_by_xpath('//*[@id="__next"]/div/div[3]/div[2]/div/div[1]/button').click()
            driver.find_element_by_xpath('//*[@id="__next"]/div/div[3]/div[2]/div/div[1]/ul/li[2]/button').click()
            sleep(0.3)
                
            SCROLL_PAUSE_TIME = 0.1
            # Get scroll height

            height = 540
            for i in range(0,20):
                driver.execute_script("window.scrollTo("+ str(height * i) +", "+ str(height * i) +");")
                sleep(SCROLL_PAUSE_TIME)
        except:
            print(data[n][0] + "제품이 없습니다.")
            cannot_get.append(data[n][0])


        html = driver.page_source

        soup = BeautifulSoup(html, 'html.parser')

        #크롤링한 html파일에 가격정보가 몇개가 들어있는지 파악
        product_review_infos = soup.find_all('li', class_='reviewItem_review_item__3Xxc9')
        product_review_len = len(product_review_infos)


        # print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>___" + data[n][0] + "___<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<" )

        #크롤링한 정보를 담을 data_arr이라는 배열,
        #한 정보의 싸이클이 끝나고 다시 시작하면 초기화.


        data_arr = []

        #product_review_infos 
        for i in range(0,product_review_len):

            metadata = product_review_infos[i]
            contents = metadata.find('span',{'class':'reviewItem_content__23gBj'}).text

            rate = metadata.find('span',{'class':'reviewItem_info_average__3R8oQ'}).text
            rate = ''.join(filter(str.isdigit, rate))
            rate = int(rate)

            writer = metadata.find('span',{'class':'reviewItem_data__3jzmv'}).text

            date = metadata.findAll('span',{'class':'reviewItem_data__3jzmv'})[1].text
            date = date.replace(".", "")
            date = int(date)
            
            mall = metadata.find('span',{'class':'reviewItem_mall__3rFFu'}).text

            if( mall == "밀레코리아" ):
                mall = "네이버" 


            try:
                image = metadata.findAll('img')[1]['src']
            except:
                image = "NO IMG"

            # print(image['src'])

            data_arr.append([contents, rate, writer, date, mall, image]) 
            # print(data_arr)



        # data_arr에는 한 제품에 대한 모든 쇼핑물의 정보가 들어있으니깐, 쇼핑물 하나하나 엑셀에다가 정리하자
        # 0 상품평
        # 1 별점 수
        # 2 작성자
        # 3 작성 날짜
        # 4 몰 이름
        # 5 이미지 url
        # data 0 상품명, 1 원부코드
        # worksheet.write(row, 0, data[n][0].upper(), title_cell) #상품명

        # for i in range(1,7):
        #     worksheet.write(row, i, " ", title_cell) #상품명
        # row = row + 1
        

        for i in range(0, len(data_arr)):

            if(data_arr[i][3] >= START_DATE and data_arr[i][3] <= END_DATE):
                # print(data_arr[i][0])

                # worksheet.set_column(col + 1, cell_format=data_format2)
                worksheet.write(row, 0, data_arr[i][4]) #쇼핑물
                worksheet.write(row, 1, data[n][0].upper()) #상품명
                worksheet.write(row, 2, data_arr[i][2]) #작성자
                worksheet.write(row, 3, data_arr[i][1]) #별점
                worksheet.write(row, 4, data_arr[i][0]) #상품평
                worksheet.write(row, 5, data_arr[i][5]) #사진 URL
                worksheet.write(row, 6, data_arr[i][3]) #작성날짜
                

                # col = col + 1
                row = row + 1



    #3번쨰 컬럼 서식 (differ 값 구하는 식)
    # for row_num in range(1, 100):
    #     worksheet.write_formula(row_num-1, 3,
    #                             '= -$C%d + $E%d' % (row_num, row_num))

    workbook.close()
    driver.close()

    return cannot_get

data = get_review("201101","20210302",self)

print(data)


