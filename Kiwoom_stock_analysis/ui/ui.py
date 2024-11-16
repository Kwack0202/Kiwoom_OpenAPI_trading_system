from kiwoom.kiwoom import *

import sys
from PyQt5.QtWidgets import * # PyQt5는 Ui를 꾸미기기 좋은 모듈

class Ui_class():
    def __init__(self):
        print("Ui_class 입니다")
        
        self.app = QApplication(sys.argv) # QApplication는 빈 깡통 상태의 어플 Ui
        
        self.kiwoom = Kiwoom()
        
        self.app.exec_() # 개장시간 동안 프로그램이 계속 돌아가기 위해서 해당 코드를 통해 py파일이 종료되지 않도록 해주는 코드