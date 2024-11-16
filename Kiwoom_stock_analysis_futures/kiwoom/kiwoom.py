from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *

import datetime
import pickle
import numpy as np
import pandas as pd
import os

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__() # super == 부모의 init함수를 사용하겠다
        
        print("Kiwoom 클래스 시작합니다.")
        
        ##### 이벤트 루프 모음 #####
        self.login_event_loop = None
        self.detail_account_info_event_loop= QEventLoop()
        self.calculator_event_loop = QEventLoop()
        
        ##### 스크린 번호 모음 #####
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"
        
        ##### 종목 정보용 초기 빈 딕셔너리 모음 #####
        self.account_stock_dict = {}
        
        self.portfolio_stock_dict = {}
        
        ##### 종목 분석 용 초기 빈 리스트 모음 ##### 
        self.calcul_kospi_data = []
        self.calcul_kosdaq_data = []
        self.day_stock = []
        
        ##### 계좌 정보 관련 변수 #####
        self.account_num = None # 계좌번호
        self.use_money = 0 #실제 투자에 사용할 금액
        self.use_money_percent = 0.5 #예수금에서 실제 사용할 비율

        
        ##### 지정 함수 호출 모음 #####
        self.get_ocx_instance() # OCX 방식을 파이썬에 사용할 수 있게 변환해주는 함수
        self.event_slot() # 키움과 연결하기 위한 시그널/슬롯 모음
        self.signal_login_commConnect() # 로그인 요청 시그널 포함
        self.get_account_info() # 계좌번호 정보 가져오기
        
        self.detail_account_info() # 증권 계좌 예수금 정보 가져오기
        self.detail_account_mystock() # 증권 계좌 잔고내역 정보 가져오기
        
    
        self.OPT50029_10100000() #코스피200 선물
        self.OPT50029_10600000() #코스닥150 선물
        
        self.pickle_data_save()

        #==========================================================================================================================================================
        #==========================================================================================================================================================     
    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # 레지스트리에 저장된 api 모듈 불러오기
        
        
    def event_slot(self):
        self.OnEventConnect.connect(self.login_slot) # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot) # TRdata 요청 이벤트
        
        
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
        print("\n----예탁금 및 증거금 요청 부분----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "선옵예탁금및증거금조회요청", "OPW20010", "0", self.screen_my_info) # 마지막 화면번호 1000: 잔고조회 / 2000: 실시간 데이터 조회(종목 1~100) / 2001: 실시간 데이터 조회(종목 101~200) / 3000: 주문요청 / 4000: 일봉조회
    
        self.detail_account_info_event_loop.exec_()
    
    
    def detail_account_mystock(self,sPrevNext="0"):
        print("\n----선옵 잔고 현황 정산가 기준 요청 부분----")
        
        self.dynamicCall("SetInputValue(String, String)", "계좌번호", self.account_num) # 계좌번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호", 0000) # 비밀번호
        self.dynamicCall("SetInputValue(String, String)", "비밀번호입력매체구분", 00)
        
        self.dynamicCall("CommRqData(String, String, int, String)", "선옵잔고현황정산가기준요청", "opw20007", sPrevNext, self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
    
    
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        -TR요청을 하는 slot-
        
        sScrNo: 스크린 번호
        sRQName: 내가 요청할 때 지은 이름
        sTrCode: 요청한 TR Code
        sRecordName: 사용 안함
        sPrevNext: 다음 페이지가 있는지 알려줌
        
        '''
        
        if sRQName == "선옵예탁금및증거금조회요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예탁총액")
            print("예탁총액 : %s" % int(deposit), "원")
            
            
            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "인출가능총액")
            print("인출가능총액: %s" % int(ok_deposit), "원")
            
            self.detail_account_info_event_loop.exit()
        #====================================================================================================================
        #====================================================================================================================   
        elif sRQName == "선옵잔고현황정산가기준요청":
            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "약정금액합계")
            print("약정금액합계 : %s" % int(total_buy_money), "원")
            
            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "평가손익합계")
            print("평가손익합계 : %s" % int(total_profit_loss_rate), "원")
            
            self.detail_account_info_event_loop.exit()
            
        #====================================================================================================================
        #====================================================================================================================
        elif sRQName == "선물옵션_코스피200_분차트요청_10100000":
            print("KOSPI200(code : 10100000) 분봉데이터 요청") 
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 900틱까지 데이터를 받을 수 있음
            print("조회 데이터 분봉 수 %s" % cnt)
                    
            for i in range(cnt):
                data = [] # 종목별.. 1일마다 빈 리스트로 만들어짐..
                
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_time = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결시간")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")
                
                # data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_time.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                # data.append("")
                
                self.calcul_kospi_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_dacalcul_kospi_datata에 append 
                
            if sPrevNext == "2":
                print("==>> 다음 페이지 데이터 조회 ==>>\n")
                self.OPT50029_10100000(sPrevNext = sPrevNext)
            else:
                print("Futures Options Code : [ 10100000 ] 수집 완료 총 %s" % len(self.calcul_kospi_data))
                
                # 데이터 저장하는 부분
                path_kospi_10100000 = 'C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/saved_options_time_data_kospi200'
                
                if not os.path.exists(path_kospi_10100000):
                    os.makedirs(path_kospi_10100000)
                kiwoom_day = os.path.join(path_kospi_10100000, '10100000_data.pkl') 
                
                if os.path.exists(kiwoom_day):
                    with open(kiwoom_day, 'rb') as f:
                        existing_data = pickle.load(f)

                    today = datetime.datetime.today().strftime('%Y%m%d')
                    todays_data = [item for item in self.calcul_kospi_data if item[2].startswith(today)]  # Assuming trading_time is in index 2

                    if todays_data:
                        updated_data = [item for item in existing_data if not item[2].startswith(today)]
                        updated_data.extend(todays_data)

                        with open(kiwoom_day, 'wb') as f:
                            pickle.dump(updated_data, f)

                        print("Data for {} updated in {}".format(today, kiwoom_day))
                    else:
                        print("No new data for today")
                else:
                    with open(kiwoom_day, 'wb') as f:
                        pickle.dump(self.calcul_kospi_data, f)
                    print("Initial data saved in {}".format(kiwoom_day))

                self.calcul_kospi_data.clear()
                self.calculator_event_loop.exit()
                
        #====================================================================================================================
        #====================================================================================================================
        elif sRQName == "선물옵션_코스닥150_분차트요청_10600000":
            print("KOSDAQ150(code : 10600000) 분봉데이터 요청") 
            
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName) # 1회 조회당 900틱까지 데이터를 받을 수 있음
            print("조회 데이터 분봉 수 %s" % cnt)
                    
            for i in range(cnt):
                data = [] # 종목별.. 1일마다 빈 리스트로 만들어짐..
                
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_time = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결시간")
                start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")
                
                # data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_time.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                # data.append("")
                
                self.calcul_kosdaq_data.append(data.copy()) # 종목별 1분씩 생성된 데이터를 self.calcul_kosdaq_data append 
                
            if sPrevNext == "2":
                print("==>> 다음 페이지 데이터 조회 ==>>\n")
                self.OPT50029_10600000(sPrevNext = sPrevNext)
            else:
                print("Futures Options Code : [ 10600000 ] 수집 완료 총 %s" % len(self.calcul_kosdaq_data))
                
                # 데이터 저장하는 부분
                path_kosdaq_10600000 = 'C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/saved_options_time_data_kosdaq150/'
                
                if not os.path.exists(path_kosdaq_10600000):
                    os.makedirs(path_kosdaq_10600000)
                kiwoom_day = os.path.join(path_kosdaq_10600000, '10600000_data.pkl') 
                
                if os.path.exists(kiwoom_day):
                    with open(kiwoom_day, 'rb') as f:
                        existing_data = pickle.load(f)

                    today = datetime.datetime.today().strftime('%Y%m%d')
                    todays_data = [item for item in self.calcul_kosdaq_data if item[2].startswith(today)]  # Assuming trading_time is in index 2

                    if todays_data:
                        updated_data = [item for item in existing_data if not item[2].startswith(today)]
                        updated_data.extend(todays_data)

                        with open(kiwoom_day, 'wb') as f:
                            pickle.dump(updated_data, f)

                        print("Data for {} updated in {}".format(today, kiwoom_day))
                    else:
                        print("No new data for today")
                else:
                    with open(kiwoom_day, 'wb') as f:
                        pickle.dump(self.calcul_kosdaq_data, f)
                    print("Initial data saved in {}".format(kiwoom_day))

                self.calcul_kosdaq_data.clear()
                self.calculator_event_loop.exit()
    
    
        
    def OPT50029_10100000(self, sPrevNext="0"): # 코스피 200 선물
        
        QTest.qWait(3600)
               
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", '10100000')
        self.dynamicCall("SetInputValue(QString, QString)", "시간단위", "1")
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "선물옵션_코스피200_분차트요청_10100000", "OPT50029", sPrevNext, "self.screen_calculation_stock")
        
        self.calculator_event_loop.exec_()
        
    def OPT50029_10600000(self, sPrevNext="0"): # 코스피 200 선물
        
        QTest.qWait(3600)
               
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", '10600000')
        self.dynamicCall("SetInputValue(QString, QString)", "시간단위", "1")
            
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "선물옵션_코스닥150_분차트요청_10600000", "OPT50029", sPrevNext, "self.screen_calculation_stock")
        
        self.calculator_event_loop.exec_()
    
    
    #=======================================================================================================
    # 데이터 저장 부분
    def pickle_data_save(self):
        
        # 년-월별로 따로 저장하는 함수
        def save_to_excel_by_month(df, code):
            grouped = df.groupby(df['time'].dt.strftime('%Y-%m'))
            
            for group_name, group_df in grouped:
                
                print(f'Saving data({code}) for [  {group_name}  ]')
                
                if not os.path.exists(f'C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/{code}'):
                    os.makedirs(f'C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/{code}')                    
                excel_name = f'C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/{code}/data_{group_name}.xlsx'
                save_to_excel_by_day(group_df, excel_name)

        # 일별로 시트가 나눠서 저장하는 함수
        def save_to_excel_by_day(df, excel_name):
            excel_writer = pd.ExcelWriter(excel_name, engine='openpyxl')
            
            for date, group_df in df.groupby(df['time'].dt.date):
                sheet_name = date.strftime('%Y-%m-%d')
                
                if sheet_name in excel_writer.sheets:
                    continue
                
                group_df.to_excel(excel_writer, sheet_name, index=False)
                    
            excel_writer.save()
            
            
        print('\n!!! 선물데이터 저장을 시작합니다 !!!\n')

        futures = ['kospi200', 'kosdaq150']
        
        for stock in futures:
            if stock == 'kospi200':
                code = '10100000'
            else:
                code = '10600000'
            
            save_path = f'C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/saved_options_time_data_{stock}'
            if not os.path.exists(save_path):
                os.makedirs(save_path)
                    
            with open(f"C:/Kiwoom_trading/Kiwoom_stock_analysis_futures/saved_options_time_data_{stock}/{code}_data.pkl", "rb") as fr:
                stock_data = pickle.load(fr, encoding='utf-8')
                
    
            stock_data = pd.DataFrame(stock_data, columns=['close', 'volume', 'time', 'open', 'high', 'low'])
            stock_data = stock_data[['time', 'open', 'high', 'low', 'close', 'volume']]
        
            for col in stock_data.columns:
                stock_data[col] = stock_data[col].str.replace(r'[+-]', '', regex=True).astype(float)
            
            stock_data['time'] = pd.to_datetime(stock_data['time'], format='%Y%m%d%H%M%S')
            stock_data = stock_data.sort_values(by='time', ascending=True).reset_index(drop=True)
            stock_data['time'] = pd.to_datetime(stock_data['time'])
            
            stock_data['time'] = pd.to_datetime(stock_data['time'])
        
            # 년-월별로 엑셀 저장
            save_to_excel_by_month(stock_data, code)
    
        self.detail_account_info_event_loop.exec_()
        
    
