import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
import main as review

driver = webdriver.Chrome('./src/chromedriver')

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("main.ui")[0]

# ////////////////////////////////////////////////////////////////
# 시작 종료 날짜
START_DATE = 210101
END_DATE = 210203

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
# ////////////////////////////////////////////////////////////////



#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.openRAW.clicked.connect(self.openRAWFunc)
        self.run.clicked.connect(self.runFunc)
        self.clear.clicked.connect(self.clearFunc)

    def openRAWFunc(self):
        print("open")
        self.showProgress.append("open Text")

    def runFunc(self):
        print("run")
        self.showProgress.append("run")
        worksheet = review.set_form()
        self.showProgress.append("Finish Set Form")
        lines = review.get_data_from_txt()
        self.showProgress.append("Finish Get Data From txt")

        
    
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