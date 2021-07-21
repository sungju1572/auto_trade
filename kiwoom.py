import sys
import os
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from errCode import *


class Kiwoom(QAxWidget): #QAxWidget 상속받음
    def __init__(self):
        super().__init__()
        self.login_event_loop = QEventLoop() # 로그인 담당 이벤트 루프(프로그램이 종료되지 않게하는 큰 틀의 루프)
        self.get_deposit_loop = QEventLoop()  # 예수금 담당 이벤트 루프
        self.get_account_evaluation_balance_loop = QEventLoop() # 계좌 담당 이벤트 루프
        
        
        # 계좌 관련 변수
        self.account_number = None
        self.total_buy_money = 0
        self.total_evaluation_money = 0
        self.total_evaluation_profit_and_loss_money = 0
        self.total_yield = 0
        
        # 예수금 관련 변수
        self.deposit = 0
        self.withdraw_deposit = 0
        self.order_deposit = 0

        # 화면 번호
        self.screen_my_account = "1000"


        # 초기 작업
        self.create_kiwoom_instance()
        self.event_collection()  # 이벤트와 슬롯을 메모리에 먼저 생성
        self.login()
        input() #스페이스바 입력
        self.get_account_info()  # 계좌 번호만 얻어오기
        self.get_deposit_info()  # 예수금 관련된 정보 얻어오기

        self.menu()
        
        
    # COM 오브젝트 생성
    def create_kiwoom_instance(self):
        # 레지스트리에 저장된 키움 openAPI 모듈 불러오기
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") #고유식별자 가져옴
        
    # 이벤트 처리
    def event_collection(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트 (.connect()는 이벤트와 슬롯을 연결하는 역할)
        self.OnReceiveTrData.connect(self.tr_slot)  # 트랜잭션 요청 관련 이벤트

    #로그인 
    def login(self):
        self.dynamicCall("CommConnect()")  # CommConnect() 시그널 함수 호출(.dynamicCall()는 서버에 데이터를 송수신해주는 기능)
        self.login_event_loop.exec_() #exec_()를 통해 이벤트 루프 실행  (다른데이터 간섭 막기)
    
    #로그인 성공 실패여부 
    def login_slot(self, err_code):
        if err_code == 0:
            print("로그인에 성공하였습니다.")
        else:
            if err_code == -106:  # 사용자가 강제로 키움api 프로그램을 종료하였을 경우
                os.system('cls')
                print(errors(err_code)[1])
                sys.exit(0)
            os.system('cls')
            print("로그인에 실패하였습니다.")
            print("에러 내용 :", errors(err_code)[1])
            sys.exit(0)
        self.login_event_loop.exit() #exit()를 통해 이벤트 루프 종료
 
    #계좌 받아오기 
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCLIST") #전체계좌 목록 반환 
        account_number = account_list.split(';')[0] #세미콜론 기준으로 첫번째 계좌 선택
        self.account_number = account_number #멤버 변수에 할당(초기화)
      
    
    #기타 개인정보
    def menu(self):
        sel = ""
        while True:
            os.system('cls') #화면 깨끗하게 비우기
            print("1. 현재 로그인 상태 확인")
            print("2. 사용자 정보 조회")
            print("3. 예수금 조회")
            print("4. 계좌 잔고 조회")
            print("Q. 프로그램 종료")
            sel = input("=> ")

            if sel == "Q" or sel == "q": #Q누르면 종료
                sys.exit(0)

            if sel == "1":
                self.print_login_connect_state()
            elif sel == "2":
                self.print_my_info()
            elif sel == "3":
                self.print_get_deposit_info()
            elif sel == "4":
                self.print_get_account_evaulation_balance_info()

    #로그인 상태확인
    def print_login_connect_state(self):
        os.system('cls')
        isLogin = self.dynamicCall("GetConnectState()") #로그인 상태 반환 1,0
        if isLogin == 1:
            print("\n현재 계정은 로그인 상태입니다.")
        else:
            print("\n현재 계정은 로그아웃 상태입니다.")
        input()

    #내정보
    def print_my_info(self):
        os.system('cls')
        user_name = self.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        user_id = self.dynamicCall("GetLoginInfo(QString)", "USER_ID")
        account_count = self.dynamicCall(
            "GetLoginInfo(QString)", "ACCOUNT_CNT")

        print(f"\n이름 : {user_name}")
        print(f"ID : {user_id}")
        print(f"보유 계좌 수 : {account_count}")
        print(f"1번째 계좌번호 : {self.account_number}")
        input()

    #예수금 정보 출력
    def print_get_deposit_info(self):
        os.system('cls')
        print(f"\n예수금 : {self.deposit}원")
        print(f"출금 가능 금액 : {self.withdraw_deposit}원")
        print(f"주문 가능 금액 : {self.order_deposit}원")
        input()

    #계좌 정보 출력
    def print_get_account_evaulation_balance_info(self):
        os.system('cls')
        print("\n<싱글데이터>")
        print(f"총매입금액 : {self.total_buy_money}원")
        print(f"총평가금액 : {self.total_evaluation_money}원")
        print(f"총평가손익금액 : {self.total_evaluation_profit_and_loss_money}원")
        print(f"총수익률 : {self.total_yield}%")
        input()


    #예수금 정보 입력 받기
    def get_deposit_info(self, nPrevNext=0): #입력값 : 계좌번호, 비밀번호, 비밀번호매체구분, 조회구분
        self.dynamicCall("SetInputValue(QString, QString)", #SetInputValue()안에 넣어서 서버에 데이터 송신
                         "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", " ")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2") #2는 모든 페이지
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "예수금상세현황요청", "opw00001", nPrevNext, self.screen_my_account) #opw00001 : 예수금 관련 tr 

        self.get_deposit_loop.exec_()

    #계좌 정보 입력 받기(이벤트 루프)
    def get_account_evaluation_balance(self, nPrevNext=0):
        self.dynamicCall("SetInputValue(QString, QString)",
                         "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", " ")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "계좌평가잔고내역요청", "opw00018", nPrevNext, self.screen_my_account)

        self.get_account_evaluation_balance_loop.exec_()

    #트렌젝션 이벤트
    def tr_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)

            withdraw_deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.withdraw_deposit = int(withdraw_deposit)

            order_deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "주문가능금액")
            self.order_deposit = int(order_deposit)
            self.cancel_screen_number(self.screen_my_account)
            self.get_deposit_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
            self.total_sell_money = int(total_buy_money)

            total_evaluation_money = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가금액")
            self.total_evaluation_money = int(total_evaluation_money)

            total_evaluation_profit_and_loss_money = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.total_evaluation_profit_and_loss_money = int(
                total_evaluation_profit_and_loss_money)

            total_yield = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총수익률(%)")
            self.total_yield = float(total_yield)
            self.cancel_screen_number(self.screen_my_account)
            self.get_account_evaluation_balance_loop.exit()

    def cancel_screen_number(self, sScrNo):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)