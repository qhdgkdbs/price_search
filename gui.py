import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
import main as review
import os


#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("main.ui")[0]



#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.openRAW.clicked.connect(self.openRAWFunc)
        self.run.clicked.connect(self.runFunc)
        self.clear.clicked.connect(self.clearFunc)
        self.openFolder.clicked.connect(self.openFolderFunc)

        self.currentDate = QDate.currentDate()
        self.END_DATE.setDate(self.currentDate)

    def openFolderFunc(self):
        print("open folder!!!")
        self.showProgress.append("폴더 엽니다!!!")
        path = "./review"
        path = os.path.realpath(path)
        os.startfile(path)

    def openRAWFunc(self):
        print("open")
        self.showProgress.append("텍스트 파일을 엽니다!!!")
        os.system("start ./src/pd_info.txt")


    def runFunc(self):
        START_DATE = self.START_DATE.date().toString("yyyyMMdd")
        END_DATE = self.END_DATE.date().toString("yyyyMMdd")

        print(START_DATE)
        print(END_DATE)

        if(START_DATE > END_DATE and START_DATE == END_DATE):
            print("시작 날짜가 마지막 날짜보다 같거나 작을 수 없습니다.")
            print("5초 후에 프로그램이 종료 됩니다.")
            sleep(5)
            sys.exit()

        try:
            START_DATE = int(START_DATE)
            END_DATE = int(END_DATE)
        except:
            print("날짜 값이 이상합니다(숫자만 입력하세요).")
            print("5초 후에 프로그램이 종료 됩니다.")
            sleep(5)
            sys.exit()

        print("run")
        self.showProgress.append("실행합니다. 기다려주세요!")
        no_data = review.get_review(START_DATE, END_DATE, self)
        self.showProgress.append("finish")
        for data in no_data:
            self.showProgress.append(data + "상품은 정보를 가져올 수 없습니다.")
        
    
    def clearFunc(self):
        print("clear")
        self.showProgress.clear()


if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()