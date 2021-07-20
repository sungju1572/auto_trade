from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from errCode import *


class Kiwoom(QAxWidget): #QAxWidget 상속받음
    def __init__(self):
        super().__init__()
        self.login_event_loop = QEventLoop() # 로그인 담당 이벤트 루프(프로그램이 종료되지 않게하는 큰 틀의 루프)

        # 계좌 관련 변수
        self.account_number = None

        # 초기 작업
        self.create_kiwoom_instance()
        self.login()
        self.get_account_info()

    # COM 오브젝트 생성
    def create_kiwoom_instance(self):
        # 레지스트리에 저장된 키움 openAPI 모듈 불러오기
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") #고유식별자 가져옴

    #로그인 
    def login(self):
        self.OnEventConnect.connect(self.login_slot)  # 이벤트와 슬롯을 메모리에 먼저 생성.(.connect()는 이벤트와 슬롯을 연결하는 역할)
        self.dynamicCall("CommConnect()")  # CommConnect() 시그널 함수 호출(.dynamicCall()는 서버에 데이터를 송수신해주는 기능)
        self.login_event_loop.exec_() #exec_()를 통해 이벤트 루프 실행  (다른데이터 간섭 막기)
    
    #로그인 성공 실패여부 
    def login_slot(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print("로그인 실패 - 에러 내용 :", errors(err_code)[1])
        self.login_event_loop.exit() #exit()를 통해 이벤트 루프 종료
 
    #계좌 받아오기 
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCLIST") #전체계좌 목록 반환 
        account_number = account_list.split(';')[0] #세미콜론 기준으로 첫번째 계좌 선택
        self.account_number = account_number #멤버 변수에 할당(초기화)
        print(self.account_number)       
    
        
        