import os
import sys

from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *
from config.kiwoomType import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__() # super == 부모의 init함수를 사용하겠다
        
        self.realType = RealType()
        print("Kiwoom 클래스 시작합니다.")
        
        ##### 이벤트 루프 모음 #####
        self.login_event_loop = None
        self.detail_account_info_event_loop= QEventLoop()
        self.calculator_event_loop = QEventLoop()
        
        ##### 스크린 번호 모음 #####
        self.screen_my_info = "2000"           # 계좌 관련 스크린 번호
        self.screen_calculation_stock = "4000" # 계산용 스크린 번호
        self.screen_real_stock = "5000"        # 종목 별 할당할 스크린 번호
        self.screen_meme_stock = "6000"        # 종목 별 할당할 주문용 스크린 번호
        self.screen_start_stop_real = "1000"   # 장 시작/종료 실시간 스크린 번호
        
        ##### 종목 정보용 초기 빈 딕셔너리 모음 #####
        self.account_stock_dict = {}
        self.not_account_stock_dict = {}
        
        self.portfolio_stock_dict = {}
        self.jango_dict = {}
        
        ##### 종목 분석 용 초기 빈 리스트 모음 ##### 
        self.calcul_data = []
        
        ##### 계좌 정보 관련 변수 #####
        self.account_num = None # 계좌번호
        self.deposit = 0 #예수금
        self.output_deposit = 0 #출력가능 금액
        
        self.use_money = 0 #실제 투자에 사용할 금액
        self.use_money_percent = 0.5 #예수금에서 실제 사용할 비율

        self.total_profit_loss_money = 0 #총평가손익금액
        self.total_profit_loss_rate = 0.0 #총수익률(%)
        
        ##### 지정 함수 호출 모음 #####
        self.get_ocx_instance() # OCX 방식을 파이썬에 사용할 수 있게 변환해주는 함수
        self.event_slot() # 키움과 연결하기 위한 시그널/슬롯 모음
        self.real_event_slot() # 키움과 연결하기 위한 실시간 시그널/슬롯 모음
        self.signal_login_commConnect() # 로그인 요청 시그널 포함
        self.get_account_info() # 계좌번호 정보 가져오기
        
        self.detail_account_info() # 증권 계좌 예수금 정보 가져오기
        self.detail_account_mystock() # 증권 계좌 잔고내역 정보 가져오기
        self.not_concluded_account() #미체결 요청
        
        self.read_code() # 저장된 종목 불러오기
        self.screen_number_setting() #스크린 번호 할당
        
        # self.OPT10023() # 거래량 급증 종목 TR 조회
        #==========================================================================================================================================================
        #==========================================================================================================================================================
        
        # 실시간 수신 관련 함수
        self.dynamicCall("SetRealReg(Qstring, Qstring, Qstring, Qstring)", self.screen_start_stop_real, "", self.realType.REALTYPE['장시작시간']['장운영구분'], "0")
        
        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]["스크린번호"]
            fids = self.realType.REALTYPE['주식체결']['체결시간']

            self.dynamicCall("SetRealReg(Qstring, Qstring, Qstring, Qstring)", screen_num, code, fids, "1")
            print("실시간 등록 코드: %s, 스크린번호: %s, fid번호: %s" % (code, screen_num, fids))
            print("===========================================================================")

        
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # 레지스트리에 저장된 api 모듈 불러오기
        
        
    def event_slot(self):
        self.OnEventConnect.connect(self.login_slot) # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot) # TRdata 요청 이벤트
        self.OnReceiveMsg.connect(self.msg_slot) 
        
    
    def real_event_slot(self):
        self.OnReceiveRealData.connect(self.realdata_slot) # 실시간 이벤트 연결
        self.OnReceiveChejanData.connect(self.chejan_slot) # 종목 주문체결 관련 이벤트
        
        
    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()") # 로그인 요청 시그널
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
        
        
    def login_slot(self, errCode):
        print(errors(errCode))
        
        self.login_event_loop.exit()
    
    
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        
        self.account_num = account_list.split(';')[1] # 0: 선물옵션 계좌번호 / 1번 8042899611 : 모의투자 상시용 계좌 / 2번 8042899711 : 모의투자 대회용 계좌
        
        print("\n나의 보유 계좌번호: %s" % self.account_num)
        
        
    def detail_account_info(self):
        print("\n----예수금 요청 부분----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        self.dynamicCall("SetInputValue(String, String)", "조회구분", 2)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "예수금상세현황요청", "opw00001", "0", self.screen_my_info) # 마지막 화면번호 1000: 잔고조회 / 2000: 실시간 데이터 조회(종목 1~100) / 2001: 실시간 데이터 조회(종목 101~200) / 3000: 주문요청 / 4000: 일봉조회
    
        self.detail_account_info_event_loop.exec_()
    
    
    def detail_account_mystock(self,sPrevNext="0"):
        print("\n----계좌평가 잔고내역 요청 연속조회----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        self.dynamicCall("SetInputValue(String, String)", "조회구분", 2)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
    
    
    def not_concluded_account(self, sPrevNext="0"):
        print("\n----미체결 요청----")
        print("===========================================================================")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "체결구분", "1") 
        self.dynamicCall("SetInputValue(String, String)", "매매구분", "0")
        
        self.dynamicCall("CommRqData(String, String, int, String)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
        print("                         미체결 요청 조회 완료!!                          ")
        print("===========================================================================\n")
        
    
    
    #=======================================================================================================================
    # 실시간으로 요청하는 정보가 아닌 과거에 저장된 정보(TR Data) 요청부문
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        -TR요청을 하는 slot-
        
        sScrNo: 스크린 번호
        sRQName: 내가 요청할 때 지은 이름
        sTrCode: 요청한 TR Code
        sRecordName: 사용 안함
        sPrevNext: 다음 페이지가 있는지 알려줌
        
        '''
        
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print("예수금 : %s" % int(deposit), "원")
            
            self.use_money = int(deposit) * self.use_money_percent
            self.use_money = self.use_money / 4
            
            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "출금가능금액")
            print("출금가능금액: %s" % int(ok_deposit), "원")
            
            self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================   
        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            print("총매입금액 : %s" % int(total_buy_money), "원")
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            print("총수익률 : %s" % float(total_profit_loss_rate), "%")
            
            
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]
                
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                
                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}
                
                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())
                
                self.account_stock_dict[code].update({"종목명" : code_nm})
                self.account_stock_dict[code].update({"보유수량" : stock_quantity})
                self.account_stock_dict[code].update({"매입가" : buy_price})
                self.account_stock_dict[code].update({"수익률(%)" : learn_rate})
                self.account_stock_dict[code].update({"현재가" : current_price})
                self.account_stock_dict[code].update({"매입금액" : total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량" : possible_quantity})
                
                cnt += 1
            
            print("계좌 보유 종목 개수: %s" % cnt, "개")
            print("계좌 보유 종목: %s" % self.account_stock_dict)
            
            
            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================
        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
    
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")            
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분")
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")
                
                
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())
                
                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}
                
                self.not_account_stock_dict[order_no].update({"종목코드": code})
                self.not_account_stock_dict[order_no].update({"종목명": code_nm})
                self.not_account_stock_dict[order_no].update({"주문번호": order_no})
                self.not_account_stock_dict[order_no].update({"주문상태": order_status})
                self.not_account_stock_dict[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dict[order_no].update({"주문가격": order_price})
                self.not_account_stock_dict[order_no].update({"주문구분": order_gubun})
                self.not_account_stock_dict[order_no].update({"미체결수량": not_quantity})
                self.not_account_stock_dict[order_no].update({"체결량": ok_quantity})
                
                print("미체결 종목 : %s" % self.not_account_stock_dict[order_no])
                     
            self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================        
        elif sRQName == "거래량급증요청":
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 600틱까지 데이터를 받을 수 있음
            print("====  Data updating ... ====")
            print("거래량 급증 종목 조회 :  데이터 종목 수 %s" % cnt)
            
            for i in range(cnt):
                data = [] # 종목별.. 1일마다 빈 리스트로 만들어짐..
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                code_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                today_updown = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "전일대비")
                updown_ratio = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "등락률")
                before_volume = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "이전거래량")
                now_volume = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재거래량")
                increase_ratio = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "급증률") 
                
                # data.append("")
                data.append(stock_code.strip())
                data.append(code_name.strip())
                data.append(int(current_price.strip()))
                data.append(int(today_updown.strip()))
                data.append(float(updown_ratio.strip()))
                data.append(int(before_volume.strip()))
                data.append(int(now_volume.strip()))
                data.append(float(increase_ratio.strip()))
                # data.append("")
                self.calcul_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_data에 append 
                
                # 종목 선정 하기
                pass_success = False
            
                if float(updown_ratio.strip()) < 0: # 등락률 0% 이상
                    pass_success = False
                    
                elif float(updown_ratio.strip()) > 20: # 등락률 15% 이하
                    pass_success = False
                
                elif int(before_volume.strip()) < 1500000: # 이전거래량 1백만 이상
                    pass_success = False
                        
                elif int(now_volume.strip()) < 1500000: # 현재거래량 1백만 이상
                    pass_success = False

                else:
                    pass_success = True
            

                if pass_success == True:
                    print("종목코드 [%s] 조건부 통과됨" % stock_code.strip())

                    code_nm = self.dynamicCall("GetMasterCodeName(QString)", stock_code.strip())

                    f = open("./files/condition_stock.txt", "a", encoding="utf8")
                    f.write("%s\t%s\t%s\n" % (stock_code.strip(), code_nm, abs(int(current_price.strip()))))
                    f.close()

                elif pass_success == False:
                    print("종목코드 [%s] 조건부 통과실패" % stock_code.strip())
                
            if sPrevNext == "2":
                print("==>> 다음 페이지 데이터 조회 ==>>\n")
                self.OPT10023(sPrevNext = sPrevNext)
            else:
                print("Stock Code : [ 총 %s 개 ] 수집 완료" % len(self.calcul_data))
                    
                self.calcul_data.clear() # 특정 종목 모든 일수 조회결과 저장 후 초기화
                self.calculator_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================
    def OPT10023(self, sPrevNext="0"):
        
        QTest.qWait(3600)
               
        self.dynamicCall("SetInputValue(QString, QString)", "시장구분", "000")
        self.dynamicCall("SetInputValue(QString, QString)", "정렬구분", "2") #급증률
        self.dynamicCall("SetInputValue(QString, QString)", "시간구분", "2") # 전일 대비
        self.dynamicCall("SetInputValue(QString, QString)", "거래량구분", "5") # 5천주 이상 
        self.dynamicCall("SetInputValue(QString, QString)", "시간", "분")
        self.dynamicCall("SetInputValue(QString, QString)", "종목조건", "1") # 관리종목 제외
        self.dynamicCall("SetInputValue(QString, QString)", "가격구분", "0")
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "거래량급증요청", "OPT10023", sPrevNext, "self.screen_calculation_stock")
        
        self.calculator_event_loop.exec_()

    
    def read_code(self): # tr요청 부분을 통해 작성된 condition_stock.txt(전략에 맞는 종목 포트폴리오)를 읽어와 portfolio_stock_dict변수에 코드 저장
        
        if os.path.exists("./files/condition_stock.txt"):
            f = open("./files/condition_stock.txt","r",encoding = "utf8")
            lines = f.readlines()
            print("\n----투자 할 포트폴리오 조회----")
            print("===========================================================================")
            for line in lines:
                if line != "":
                    ls = line.split("\t")
                    
                    stock_code = ls[0]
                    stock_name = ls[1]
                    stock_price = int(ls[2].split("\n")[0])
                    stock_price = abs(stock_price)
                    
                    self.portfolio_stock_dict.update({stock_code:{"종목명":stock_name, "현재가":stock_price}})
            f.close()
            print(self.portfolio_stock_dict)
            print("                     투자 할 포트폴리오 조회 완료!!                         ")
            print("===========================================================================\n")
    
    def screen_number_setting(self):
        
        screen_overwrite = []
        
        # 계좌평가잔고내역에 있는 종목들
        for code in self.account_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)
        
        # 미체결에 있는 종목들
        for order_number in self.not_account_stock_dict.keys():
            code = self.not_account_stock_dict[order_number]["종목코드"]
            
            if code not in screen_overwrite:
                screen_overwrite.append(code)
        
        # 포트폴리오에 담겨있는 종목들
        for code in self.portfolio_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)
                
        ####################################
        # 스크린 번호 할당
        cnt = 0
        for code in screen_overwrite:
            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)
            
            if (cnt % 50) == 0:
                temp_screen += 1
                self.screen_real_stock = str(temp_screen)
            
            if (cnt % 50) == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)
                
            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({"스크린번호":str(self.screen_real_stock)})
                self.portfolio_stock_dict[code].update({"주문용스크린번호":str(self.screen_meme_stock)})
            
            elif code not in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict.update({code:{"스크린번호":str(self.screen_real_stock), "주문용스크린번호":str(self.screen_meme_stock)}})
                
            cnt += 1
        # print(self.portfolio_stock_dict)
        
        
    def realdata_slot(self, sCode, sRealType, sRealData):
        
        if sRealType == "장시작시간":
            fid = self.realType.REALTYPE[sRealType]["장운영구분"]
            value = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, fid)
            
            if value == "0":
                print("장 시작 전")
            
            elif value == "3":
                print("장 시작 9:00")
                
            elif value == "2":
                print("장 마감 전 동시호가 진행 15:20 ~ 15:30")
            
            elif value == "4":
                print("장 종료 - <예상지수 종료>")
                
                for code in self.portfolio_stock_dict.keys():
                    self.dynamicCall("SetRealRemove(QString, QString)", self.portfolio_stock_dict[code]['스크린번호'], code)
                
                QTest.qWait(5000)
                # self.file_delete()
                # self.OPT10023()
                sys.exit()
                
        
        elif sRealType == "주식체결":
            a = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, self.realType.REALTYPE[sRealType]['체결시간'])
            
            b = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, self.realType.REALTYPE[sRealType]['현재가'])
            b = abs(int(b))
            
            c = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, self.realType.REALTYPE[sRealType]['전일대비'])
            c = int(c)
            
            d = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, self.realType.REALTYPE[sRealType]['등락율'])
            d = float(d)
            
            e = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, self.realType.REALTYPE[sRealType]['(최우선)매도호가'])
            e = abs(int(e))
            
            f = self.dynamicCall("GetCommRealData(Qstring, int)", sCode, self.realType.REALTYPE[sRealType]['(최우선)매수호가'])
            f = abs(int(f))
            
            g = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['거래량'])  
            g = abs(int(g))
            
            h = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['누적거래량'])  
            h = abs(int(h))

            i = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['고가'])  
            i = abs(int(i))

            j = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['시가'])  
            j = abs(int(j))

            k = self.dynamicCall("GetCommRealData(QString, int)", sCode, self.realType.REALTYPE[sRealType]['저가']) 
            k = abs(int(k))
            
            if sCode not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.update({sCode:{}})

            self.portfolio_stock_dict[sCode].update({"체결시간": a})
            self.portfolio_stock_dict[sCode].update({"현재가": b})
            self.portfolio_stock_dict[sCode].update({"전일대비": c})
            self.portfolio_stock_dict[sCode].update({"등락율": d})
            self.portfolio_stock_dict[sCode].update({"(최우선)매도호가": e})
            self.portfolio_stock_dict[sCode].update({"(최우선)매수호가": f})
            self.portfolio_stock_dict[sCode].update({"거래량": g})
            self.portfolio_stock_dict[sCode].update({"누적거래량": h})
            self.portfolio_stock_dict[sCode].update({"고가": i})
            self.portfolio_stock_dict[sCode].update({"시가": j})
            self.portfolio_stock_dict[sCode].update({"저가": k})
            
            print(self.portfolio_stock_dict[sCode])
            
            #====================================================================================================================
            # 계좌잔고평가내역에 있고 오늘 산 잔고내역에는 없는 경우
            if sCode in self.account_stock_dict.keys() and sCode not in self.jango_dict.keys():
                print("%s %s" % ("신규매도를 한다", sCode))
                
                asd = self.account_stock_dict[sCode]
                meme_rate = (b - asd['매입가']) / asd['매입가'] * 100
                
                if asd['매매가능수량'] > 0 and (meme_rate > 7 or meme_rate < -7): # 특정 종목 누적 수익률 7% 이상 or 수익률 -7% 이하
                    order_success = self.dynamicCall("SendOrder(Qstring, Qstring, Qstring, int, Qstring, int, int, Qstring, Qstring)",
                                                     ["신규매도", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 2, sCode, asd['매매가능수량'], 0, self.realType.SENDTYPE['거래구분']['시장가'], ""])
                    if order_success == 0:
                        print("매도주문 전달 성공")
                        del self.account_stock_dict[sCode]
                    else:
                        print("매도주문 전달 실패")
            
            #====================================================================================================================    
            # 계좌잔고평가내역에 있고 오늘 산 잔고내역에 있는 경우
            elif sCode in self.jango_dict.keys():
                print("%s %s" % ("신규매도를 한다2", sCode))
                
                jd = self.jango_dict[sCode]
                meme_rate = (b - jd['매입단가']) / jd['매입단가'] * 100

                if jd['주문가능수량'] > 0 and (meme_rate > 7 or meme_rate < -7): # 수익률 5% 이상 or 수익률 -5% 이하

                    order_success = self.dynamicCall(
                        "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                        ["신규매도", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 2, sCode, jd['주문가능수량'], 0, self.realType.SENDTYPE['거래구분']['시장가'], ""]
                    )

                    if order_success == 0:
                        print("매도주문 전달 성공")
                    else:
                        print("매도주문 전달 실패")
                        
            #====================================================================================================================
            # 등락율이 2.0% 이상이고 오늘 산 잔고내역에는 없는 경우
            elif d > 2.0 and sCode not in self.jango_dict.keys():
                print("%s %s" % ("신규매수를 한다", sCode))
                
                print("매수조건 통과 %s " % sCode)

                if e != 0:
                    result = (self.use_money * 0.05) / e
                    quantity = int(result)
                    order_success = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                                     ["신규매수", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 1, sCode, quantity, e, self.realType.SENDTYPE['거래구분']['지정가'], ""])
                else:
                    order_success = 1 
                
                
                if order_success == 0:
                    print("매수주문 전달 성공")
                else:
                    print("매수주문 전달 실패")
                
            not_meme_list = list(self.not_account_stock_dict)
            for order_num in not_meme_list:
                code = self.not_account_stock_dict[order_num]["종목코드"]
                meme_price = self.not_account_stock_dict[order_num]['주문가격']
                not_quantity = self.not_account_stock_dict[order_num]['미체결수량']
                order_gubun = self.not_account_stock_dict[order_num]['주문구분']


                if order_gubun == "매수" and not_quantity > 0 and e > meme_price:
                    order_success = self.dynamicCall(
                        "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                        ["매수취소", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num, 3, code, 0, 0, self.realType.SENDTYPE['거래구분']['지정가'], order_num])

                    if order_success == 0:
                        print("매수취소 전달 성공")
                    else:
                        print("매수취소 전달 실패")

                elif not_quantity == 0:
                    del self.not_account_stock_dict[order_num]
    
    
    # 실시간 체결 정보
    def chejan_slot(self, sGubun, nItemCnt, sFIdList):
        if int(sGubun) == 0: # 주문체결
            account_num = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['종목명'])
            stock_name = stock_name.strip()

            origin_order_number = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['원주문번호']) 
            order_number = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문번호'])  

            order_status = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문상태']) 
            order_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문수량'])  
            order_quan = int(order_quan)

            order_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문가격'])  
            order_price = int(order_price)

            not_chegual_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['미체결수량'])  
            not_chegual_quan = int(not_chegual_quan)

            order_gubun = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문구분']) 
            order_gubun = order_gubun.strip().lstrip('+').lstrip('-')

            chegual_time_str = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문/체결시간']) 

            chegual_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['체결가'])  
            if chegual_price == '':
                chegual_price = 0
            else:
                chegual_price = int(chegual_price)

            chegual_quantity = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['체결량']) 
            if chegual_quantity == '':
                chegual_quantity = 0
            else:
                chegual_quantity = int(chegual_quantity)

            current_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['현재가'])
            current_price = abs(int(current_price))

            first_sell_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['(최우선)매도호가'])
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['(최우선)매수호가'])
            first_buy_price = abs(int(first_buy_price))

            ######## 새로 들어온 주문이면 주문번호 할당
            if order_number not in self.not_account_stock_dict.keys():
                self.not_account_stock_dict.update({order_number: {}})

            self.not_account_stock_dict[order_number].update({"종목코드": sCode})
            self.not_account_stock_dict[order_number].update({"주문번호": order_number})
            self.not_account_stock_dict[order_number].update({"종목명": stock_name})
            self.not_account_stock_dict[order_number].update({"주문상태": order_status})
            self.not_account_stock_dict[order_number].update({"주문수량": order_quan})
            self.not_account_stock_dict[order_number].update({"주문가격": order_price})
            self.not_account_stock_dict[order_number].update({"미체결수량": not_chegual_quan})
            self.not_account_stock_dict[order_number].update({"원주문번호": origin_order_number})
            self.not_account_stock_dict[order_number].update({"주문구분": order_gubun})
            self.not_account_stock_dict[order_number].update({"주문/체결시간": chegual_time_str})
            self.not_account_stock_dict[order_number].update({"체결가": chegual_price})
            self.not_account_stock_dict[order_number].update({"체결량": chegual_quantity})
            self.not_account_stock_dict[order_number].update({"현재가": current_price})
            self.not_account_stock_dict[order_number].update({"(최우선)매도호가": first_sell_price})
            self.not_account_stock_dict[order_number].update({"(최우선)매수호가": first_buy_price})
            
        
        elif int(sGubun) ==1: #잔고
            account_num = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['종목코드'])[1:]

            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['종목명'])
            stock_name = stock_name.strip()

            current_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['현재가'])
            current_price = abs(int(current_price))

            stock_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['보유수량'])
            stock_quan = int(stock_quan)

            like_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['주문가능수량'])
            like_quan = int(like_quan)

            buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['매입단가'])
            buy_price = abs(int(buy_price))

            total_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['총매입가']) # 계좌에 있는 종목의 총매입가
            total_buy_price = int(total_buy_price)

            meme_gubun = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['매도매수구분'])
            meme_gubun = self.realType.REALTYPE['매도수구분'][meme_gubun]

            first_sell_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['(최우선)매도호가'])
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['(최우선)매수호가'])
            first_buy_price = abs(int(first_buy_price))

            if sCode not in self.jango_dict.keys():
                self.jango_dict.update({sCode:{}})

            self.jango_dict[sCode].update({"현재가": current_price})
            self.jango_dict[sCode].update({"종목코드": sCode})
            self.jango_dict[sCode].update({"종목명": stock_name})
            self.jango_dict[sCode].update({"보유수량": stock_quan})
            self.jango_dict[sCode].update({"주문가능수량": like_quan})
            self.jango_dict[sCode].update({"매입단가": buy_price})
            self.jango_dict[sCode].update({"총매입가": total_buy_price})
            self.jango_dict[sCode].update({"매도매수구분": meme_gubun})
            self.jango_dict[sCode].update({"(최우선)매도호가": first_sell_price})
            self.jango_dict[sCode].update({"(최우선)매수호가": first_buy_price})

            if stock_quan == 0:
                del self.jango_dict[sCode]
                self.dynamicCall("SetRealRemove(Qstring, Qstring)", self.portfolio_stock_dict[sCode]["스크린번호"], sCode)
                
    # 송수신 메세지 get
    def msg_slot(self, sScrNo, sRQName, sTrCode, msg):
        print("스크린: %s, 요청이름: %s, tr코드: %s --- %s" % (sScrNo, sRQName, sTrCode, msg))
    
    #파일 삭제
    def file_delete(self):
        if os.path.isfile("./files/condition_stock.txt"):
            os.remove("./files/condition_stock.txt")