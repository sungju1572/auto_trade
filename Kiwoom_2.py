import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import pandas as pd
import sqlite3
import datetime
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

TR_REQ_TIME_INTERVAL = 0.2

class Kiwoom(QAxWidget):
    def __init__(self, ui):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.price = 0
        self.time = ""
        self.data = 0
        self.first_data = ""
        self.ui = ui
        self.account = ""
        self.code = ""
        self.state = "초기상태"
        self.sell_time = ""
        self.present_time = "" 
        self.trade_count = 0
        self.ticker = ""
        self.liquidation = ""
        self.trade_dic = {}
        self.first_price = 0
        self.first_price_list = []
        self.first_price_range = []
        self.trade_start = False
        self.trade_set = True
        self.constant_present_price = ""
        self.constant_present_price_idx_high = ""
        self.constant_present_price_idx_low = ""
        
        #그래프
        self.fig = plt.Figure()
        self.canvas = FigureCanvas(self.fig)

        
    #COM오브젝트 생성
    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") #고유 식별자 가져옴

    #이벤트 처리
    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect) # 로그인 관련 이벤트 (.connect()는 이벤트와 슬롯을 연결하는 역할)
        self.OnReceiveTrData.connect(self._receive_tr_data) # 트랜잭션 요청 관련 이벤트
        self.OnReceiveChejanData.connect(self._receive_chejan_data) #체결잔고 요청 이벤트
        self.OnReceiveRealData.connect(self._handler_real_data) #실시간 데이터 처리

    #로그인
    def comm_connect(self):
        self.dynamicCall("CommConnect()") # CommConnect() 시그널 함수 호출(.dynamicCall()는 서버에 데이터를 송수신해주는 기능)
        self.login_event_loop = QEventLoop() # 로그인 담당 이벤트 루프(프로그램이 종료되지 않게하는 큰 틀의 루프)
        self.login_event_loop.exec_() #exec_()를 통해 이벤트 루프 실행  (다른데이터 간섭 막기)

    #이벤트 연결 여부
    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    #종목리스트 반환
    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market) #종목리스트 호출
        code_list = code_list.split(';')
        return code_list[:-1]

    #종목명 반환
    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code) #종목명 호출
        return code_name

    #통신접속상태 반환
    def get_connect_state(self):
        ret = self.dynamicCall("GetConnectState()") #통신접속상태 호출
        return ret

    #로그인정보 반환
    def get_login_info(self, tag):
        ret = self.dynamicCall("GetLoginInfo(QString)", tag) #로그인정보 호출
        return ret

    #TR별 할당값 지정하기
    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value) #SetInputValue() 밸류값으로 원하는값지정 ex) SetInputValue("비밀번호"	,  "")

    #통신데이터 수신(tr)
    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no) 
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    #실제 데이터 가져오기
    def _comm_get_data(self, code, real_type, field_name, index, item_name): 
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)", code, #더이상 지원 안함??
                               real_type, field_name, index, item_name)
        return ret.strip()

    #수신받은 데이터 반복횟수
    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    #주문 (주식)
    def send_order(self, rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no):
        self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                         [rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no])
        
    #주문 (선물)    
    def send_order_fo(self, rqname, screen_no, acc_no,  code, order_type, slbytp, hoga, quantity, price, order_no):
        self.dynamicCall("SendOrderFO(QString, QString, QString, QString, int, QString, QString, int, QString, QString)",
                         [rqname, screen_no, acc_no, code, order_type, slbytp, hoga, quantity, price, order_no])


####
    #실시간 조회관련 핸들
    def _handler_real_data(self, trcode, ret):
        
        #ui에서 계좌랑 종목코드 가져오기
        self.account = self.ui.comboBox.currentText()
        self.code = self.ui.lineEdit.text()
    
        
        #체결시간
        self.time =  self.get_comm_real_data(trcode, 20)
        
        #현재시간
        now = datetime.datetime.now()
        date = now.strftime("%Y-%m-%d ")        
        hour = now.hour
        minute = now.minute
        self.present_time = now.strftime("%H : %M")
        
        #티커와 강제청산 기준
        self.ticker = float(self.ui.lineEdit_4.text())
        self.liquidation = float(self.ui.lineEdit_5.text())
        
        #거래 시작할 타이밍(초기값 False)
        self.trade_start = self.ui.trade_set

    
        #강제청산할 시간 ui에서 가져오기
        self.sell_time = int(self.ui.comboBox_7.currentText())    
        
        
        if self.time != "":
            self.time =  datetime.datetime.strptime(date + self.time, "%Y-%m-%d %H%M%S")

        #기준시간 입력
        standard_time = float(self.ui.lineEdit_6.text())


        #초기 기준값 입력받기 (여기서는 일단입력받고 strategy_2를통해 갱신)
        if self.first_data == "":
            self.first_data = self.get_comm_real_data(trcode, 10)
            if self.first_data != "":
                self.first_data = float(self.first_data[1:])
                print("기준가격:" , self.first_data, end=" ")
                print("상태: ", self.state)
                print("")


        #버튼 눌렀을때 거래시작 (버튼이 눌리고, 초기 상태(싯가기준 범위가 안만들어져있고),  현재시간이 강제 청산시간보다 이를 때 조건 )
        if self.trade_start == True and self.trade_set == True and  hour < self.sell_time:
            
            
            self.first_price_range = self.first_price_list
            
            if self.constant_present_price == "":
                self.constant_present_price  = self.get_comm_real_data(trcode, 10)
                
              
                if self.constant_present_price !="":
                    self.constant_present_price = float(self.constant_present_price[1:])
                    self.first_price_range.append(self.constant_present_price)
                    self.first_price_range = sorted(self.first_price_range)
                    self.first_price_range = list(np.round(self.first_price_range, 2))
                    idx = self.first_price_range.index(self.constant_present_price)
                    self.constant_present_price_idx_high = self.first_price_range[idx+1]
                    self.constant_present_price_idx_low = self.first_price_range[idx-1]
                    

            # 현재가 
            self.price =  self.get_comm_real_data(trcode, 10)
            self.price = self.price[1:]
            
            
            if self.price !="":
                self.price = float(self.price)
                
                
                print(self.time)
                print("|현재가: ", self.price)
                print("|초기거래 (시가기준 ):", self.trade_start )
                print("|시가 : ", self.start_price)
                print("")
                

                self.strategy_2(self.price)
                
                self.ui.present_price()
                
        # 강제청산시간이 0 혹은 현재시간이 강제청산시간보다 이르면서, 현재시간이 기준시간보다 빠르고 strategy_2함수로 인해 초기상태가 False 일 때 조건
        elif self.sell_time == 0 or hour < self.sell_time and hour >= standard_time and self.trade_set == False: 

            # 현재가 
    
            self.price =  self.get_comm_real_data(trcode, 10)
            self.price = self.price[1:]
            
            print("-----------------------------")
            print("|기준가격: " , self.first_data)
            print("|상태: ", self.state)
            print("|거래량: ", self.trade_count)
            print("|티커: ", self.ticker)
            print("초기거래 (시가기준 ):", self.trade_start )
            print("-----------------------------")
            print("")
            
            if self.price !="":
                self.price = float(self.price)
                print(self.time)
                print("현재가: ", self.price)
                print("")
                self.strategy(self.price, self.ticker, self.liquidation)
                self.ui.present_price()
        

        elif hour == self.sell_time:
            if self.state == "숏포지션":
                #청산(매수)
                self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "2", "3", 1, "0", "")
                print("청산_매수")
                print("장마무리")
                self.state = "초기상태"
            
            elif self.state == "롱포지션":
                #청산(매도)
                self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "1", "3", 1, "0", "")
                print("청산_매도")
                print("장마무리")
                self.state = "초기상태"
        else:
            print("프로그램 시작버튼을 눌러주세요")
            print("현재시간 : ", self.present_time)
            

    #실시간 데이터 가져오기
    def get_comm_real_data(self, trcode, fid):
        ret = self.dynamicCall("GetCommRealData(QString, int)", trcode, fid)
        return ret

    #체결정보
    def get_chejan_data(self, fid):
        ret = self.dynamicCall("GetChejanData(int)", fid)
        return ret
    

    def get_server_gubun(self):
        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret

    def _receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(gubun)
        print(self.get_chejan_data(9203))
        print(self.get_chejan_data(302))
        print(self.get_chejan_data(900))
        print(self.get_chejan_data(901))

    #받은 tr데이터가 무엇인지, 연속조회 할수있는지
    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2': 
            self.remained_data = True
        else:
            self.remained_data = False
            
        #받은 tr에따라 각각의 함수 호출
        if rqname == "opt10081_req":
            self._opt10081(rqname, trcode)
        elif rqname == "opw00001_req":
            self._opw00001(rqname, trcode)
        elif rqname == "opw00018_req":
            self._opw00018(rqname, trcode)
        elif rqname == "opw20006_req":
            self._opw20006(rqname, trcode)
        elif rqname == "opt50001_req":
            self._opt50001(rqname, trcode)
        elif rqname == "opt50003_req":
            self._opt50003(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    @staticmethod
    #입력받은데이터 정제    
    def change_format(data):
        strip_data = data.lstrip('-0')
        if strip_data == '' or strip_data == '.00':
            strip_data = '0'

        try:
            format_data = format(int(strip_data), ',d')
        except:
            format_data = format(float(strip_data))
        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    #입력받은데이터(수익률) 정제
    @staticmethod
    def change_format2(data):
        strip_data = data.lstrip('-0')

        if strip_data == '':
            strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    def _opw00001(self, rqname, trcode):
        d2_deposit = self._comm_get_data(trcode, "", rqname, 0, "d+2추정예수금")
        self.d2_deposit = Kiwoom.change_format(d2_deposit)


    def _opt10081(self, rqname, trcode):
        data_cnt = self._get_repeat_cnt(trcode, rqname) #데이터 갯수 확인

        for i in range(data_cnt): #시고저종 거래량 가져오기
            date = self._comm_get_data(trcode, "", rqname, i, "일자")
            open = self._comm_get_data(trcode, "", rqname, i, "시가")
            high = self._comm_get_data(trcode, "", rqname, i, "고가")
            low = self._comm_get_data(trcode, "", rqname, i, "저가")
            close = self._comm_get_data(trcode, "", rqname, i, "현재가")
            volume = self._comm_get_data(trcode, "", rqname, i, "거래량")

            self.ohlcv['date'].append(date)
            self.ohlcv['open'].append(int(open))
            self.ohlcv['high'].append(int(high))
            self.ohlcv['low'].append(int(low))
            self.ohlcv['close'].append(int(close))
            self.ohlcv['volume'].append(int(volume))
            
    #시가 가져와서 초기 리스트 만들기    
    def _opt50003(self, rqname, trcode):
        self.start_price = self._comm_get_data(trcode, "", rqname, 0, "시가")
        self.start_price = float(self.start_price[1:])
        
        ticker = float(self.ui.lineEdit_4.text())
        
        self.first_price_list = list(np.arange(self.start_price - 10, self.start_price + 10, ticker))
        
        ax = self.fig.add_subplot(111)
        
        for i in self.first_price_list:
            ax.hlines(i, 0, 1, color='gray', linestyle='-', linewidth=2)
            
    
        self.ui.graph_verticalLayout.addWidget(self.canvas)
        
        
        #self.ui.lineEdit_7.setText(str(self.first_price_list))
        
        
    #선물 현재가 가져오기(실시간x)
    def _opt50001(self, rqname, trcode):
        self.now_price = self._comm_get_data(trcode, "", rqname, 0, "현재가")
        

    #opw박스 초기화 (주식)
    def reset_opw00018_output(self):
        self.opw00018_output = {'single': [], 'multi': []}

    #여러 정보들 저장 (주식)
    def _opw00018(self, rqname, trcode):
        # single data
        total_purchase_price = self._comm_get_data(trcode, "", rqname, 0, "총매입금액")
        total_eval_price = self._comm_get_data(trcode, "", rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, 0, "총평가손익금액")
        total_earning_rate = self._comm_get_data(trcode, "", rqname, 0, "총수익률(%)")
        estimated_deposit = self._comm_get_data(trcode, "", rqname, 0, "추정예탁자산")

        self.opw00018_output['single'].append(Kiwoom.change_format(total_purchase_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_price))
        self.opw00018_output['single'].append(Kiwoom.change_format(total_eval_profit_loss_price))

        total_earning_rate = Kiwoom.change_format(total_earning_rate)
        

        if self.get_server_gubun():
            total_earning_rate = float(total_earning_rate) / 100
            total_earning_rate = str(total_earning_rate)

        self.opw00018_output['single'].append(total_earning_rate)

        self.opw00018_output['single'].append(Kiwoom.change_format(estimated_deposit))

        # multi data
        rows = self._get_repeat_cnt(trcode, rqname)
        for i in range(rows):
            name = self._comm_get_data(trcode, "", rqname, i, "종목명")
            quantity = self._comm_get_data(trcode, "", rqname, i, "보유수량")
            purchase_price = self._comm_get_data(trcode, "", rqname, i, "매입가")
            current_price = self._comm_get_data(trcode, "", rqname, i, "현재가")
            eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, i, "평가손익")
            earning_rate = self._comm_get_data(trcode, "", rqname, i, "수익률(%)")

            quantity = Kiwoom.change_format(quantity)
            purchase_price = Kiwoom.change_format(purchase_price)
            current_price = Kiwoom.change_format(current_price)
            eval_profit_loss_price = Kiwoom.change_format(eval_profit_loss_price)
            earning_rate = Kiwoom.change_format2(earning_rate)

            self.opw00018_output['multi'].append([name, quantity, purchase_price, current_price, eval_profit_loss_price,
                                                  earning_rate])

    #opw박스 초기화(선물)
    def reset_opw20006_output(self):
        self.opw20006_output = {'single': [], 'multi': []}
        
    #여러 정보들 저장 (선물)
    def _opw20006(self, rqname, trcode):
        # single data
        total_purchase_price = self._comm_get_data(trcode, "", rqname, 0, "총매입금액")
        total_eval_price = self._comm_get_data(trcode, "", rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, 0, "총평가손익금액")
        total_earning_rate = self._comm_get_data(trcode, "", rqname, 0, "총수익률(%)")
        estimated_deposit = self._comm_get_data(trcode, "", rqname, 0, "추정예탁자산")        

        self.opw20006_output['single'].append(Kiwoom.change_format(total_purchase_price))
        self.opw20006_output['single'].append(Kiwoom.change_format(total_eval_price))
        self.opw20006_output['single'].append(Kiwoom.change_format(total_eval_profit_loss_price))

        total_earning_rate = Kiwoom.change_format(total_earning_rate)

        if self.get_server_gubun():
            total_earning_rate = float(total_earning_rate) / 100
            total_earning_rate = str(total_earning_rate)

        self.opw20006_output['single'].append(total_earning_rate)

        self.opw20006_output['single'].append(Kiwoom.change_format(estimated_deposit))

        # multi data
        rows = self._get_repeat_cnt(trcode, rqname)
        for i in range(rows):
            name = self._comm_get_data(trcode, "", rqname, i, "종목명")
            quantity = self._comm_get_data(trcode, "", rqname, i, "잔고수량")
            purchase_price = self._comm_get_data(trcode, "", rqname, i, "매입단가")
            current_price = self._comm_get_data(trcode, "", rqname, i, "현재가")
            eval_profit_loss_price = self._comm_get_data(trcode, "", rqname, i, "평가손익")
            earning_rate = self._comm_get_data(trcode, "", rqname, i, "손익율")

            quantity = Kiwoom.change_format(quantity)
            purchase_price = Kiwoom.change_format(purchase_price)
            current_price = Kiwoom.change_format(current_price)
            eval_profit_loss_price = Kiwoom.change_format(eval_profit_loss_price)
            earning_rate = Kiwoom.change_format2(earning_rate)

            self.opw20006_output['multi'].append([name, quantity, purchase_price, current_price, eval_profit_loss_price,
                                                  earning_rate])
        
    
 ###
    def first_price(self):
        self.set_input_value("종목코드", self.code)
        self.comm_rq_data("opt50003_req", "opt50003", 0, "1000")
        
        
    #전략
    def strategy(self, present_price, ticker, liquidation):
       
        data = present_price
                
        
        #초기 상태
        if self.state == "초기상태":
            #매수
            if data >= self.first_data + ticker:
                self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "2", "3", 1, "0", "")
                
                print(type(data), type(self.first_data + ticker))
                
                print("매수", data )
                print("상태 : 롱포지션 진입")
                print("")
                self.first_data = data
                self.state = "롱포지션"
                self.trade_dic[self.present_time] = "롱진입"
            #매도
            elif data <= self.first_data - ticker:
                self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "1", "3", 1, "0", "")
                
                print(type(data), type(self.first_data + ticker))
                
                print("매도", data)
                print("상태 : 숏포지션 진입")
                print("")
                self.first_data = data
                self.state = "숏포지션"
                self.trade_dic[self.present_time] = "숏진입"
                
        #매수 포지션      
        elif self.state == "롱포지션":
            #매도
            if data <= self.first_data - liquidation:
                self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "1", "3", 1, "0", "")
                print("매도", data)
                print("상태 : 롱포지션 청산- /초기상태 진입")
                print("")
                self.first_data = data
                self.trade_count += 1
                self.state = "초기상태"
                self.trade_dic[self.present_time] = "롱청산"
            #윗단계로 기준 바꾸고 홀딩
            elif data >= self.first_data + ticker:
                self.first_data = data
                print("한단계위 진입", data)
                print("")
                
        #매도 포지션      
        elif self.state == "숏포지션":
            if data >= self.first_data + liquidation:
                self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "2", "3", 1, "0", "")
                print("매수", data )
                print("상태 : 숏포지션 청산- /초기상태 진입")
                print("")
                self.first_data = data
                self.trade_count += 1
                self.state = "초기상태"
                self.trade_dic[self.present_time] = "숏청산"  
            #아랫단계로 기준 바꾸고 홀딩
            elif data <= self.first_data - ticker:
                self.first_data = data
                print("한단계 아랫단계 진입", date)
                print("")
                
    #시가 기준으로 만든 범위에 따라 현재가가 그 라인에 진입시 매도 or 매수         
    def strategy_2(self, present_price):
        
        data = present_price
       # print(first_price_list)  
        print("기준점 상 : ", self.constant_present_price_idx_high)
        print("기준점 하 : ", self.constant_present_price_idx_low)              
        #초기 상태
        #매수
        if data >= self.constant_present_price_idx_high:
            self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "2", "3", 1, "0", "")
   
            print("매수", data )
            print("상태 : 롱포지션 진입")
            print("")
            self.first_data = data
            self.state = "롱포지션"
            self.trade_dic[self.present_time] = "롱진입"
            self.trade_set = False
        #매도
        elif data < self.constant_present_price_idx_low:
            self.send_order_fo("send_order_fo_req", "0101", self.account, self.code, 1, "1", "3", 1, "0", "")
                    
            print("매도", data)
            print("상태 : 숏포지션 진입")
            print("")
            self.first_data = data
            self.state = "숏포지션"
            self.trade_dic[self.present_time] = "숏진입"
            self.trade_set = False
                
          
                
    
                
    """
    #입력받은 데이터(시가) 기준으로 축만들기 
    def axis(self, first_price, ticker):
        self.first_price_list = list(np.arange(first_price -100, frist_price + 100, ticker))
        return first_price
    """       

        



if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect() #연결
    

    
    
