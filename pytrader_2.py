import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from Kiwoom_2 import *
import time

form_class = uic.loadUiType("pytrader.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
            
        self.trade_set = False
        
        self.trade_stocks_done = False

        self.kiwoom = Kiwoom(self) #객체생성
        self.kiwoom.comm_connect() #연결

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self.timer2 = QTimer(self)
        self.timer2.start(1000 *10)
        self.timer2.timeout.connect(self.timeout2)

        accouns_num = int(self.kiwoom.get_login_info("ACCOUNT_CNT"))
        accounts = self.kiwoom.get_login_info("ACCNO")

        accounts_list = accounts.split(';')[0:accouns_num]
        self.comboBox.addItems(accounts_list)

        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton.clicked.connect(self.send_order)
        self.pushButton_5.clicked.connect(self.send_order_fo)
        self.pushButton_2.clicked.connect(self.check_balance)
        self.pushButton_4.clicked.connect(self.check_balance_2)
        self.pushButton_6.clicked.connect(self.set_real_data)
        self.pushButton_7.clicked.connect(self.trade_start)
        self.pushButton_8.clicked.connect(self.start_price_list)
        

        
        

    def code_changed(self):
        code = self.lineEdit.text()
        name = self.kiwoom.get_master_code_name(code)
        self.lineEdit_2.setText(name)
        
    #계좌설정
    def set_account(self):
        account = self.comboBox.currentText()
        return account
        
##
    def set_real_data(self):
        code = self.lineEdit.text()
        self.kiwoom.set_input_value("종목코드", code)
        self.kiwoom.comm_rq_data("opt50001_req", "opt50001", 0, "1000")



    #주문 (주식)
    def send_order(self):
        order_type_lookup = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4}
        hoga_lookup = {'지정가': "00", '시장가': "03"}

        account = self.comboBox.currentText()
        order_type = self.comboBox_2.currentText()
        code = self.lineEdit.text()
        hoga = self.comboBox_3.currentText()
        num = self.spinBox.value()
        price = self.spinBox_2.value()

        self.kiwoom.send_order("send_order_req", "0101", account, order_type_lookup[order_type], code, num, price, hoga_lookup[hoga], "")

    #주문 (선물)
    def send_order_fo(self):
        order_type_lookup = {'신규매매': 1, '정정': 2, '취소': 3}
        slbytp_lookup = {"매도" : "1", "매수" : "2"}
        hoga_lookup = {'지정가': "1", '시장가': "3"}

        account = self.comboBox.currentText()
        code = self.lineEdit.text()
        order_type = self.comboBox_4.currentText()
        slbytp = self.comboBox_5.currentText()
        hoga = self.comboBox_6.currentText()
        num = self.spinBox.value()
        price = self.spinBox_2.value()

        self.kiwoom.send_order_fo("send_order_fo_req", "0101", account, code, order_type_lookup[order_type], slbytp_lookup[slbytp], hoga_lookup[hoga], num, price, "")


    #서버연결
    def timeout(self):
        market_start_time = QTime(9, 0, 0)
        current_time = QTime.currentTime()

        if current_time > market_start_time and self.trade_stocks_done is False:
            #self.trade_stocks()
            self.trade_stocks_done = True

        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = self.kiwoom.get_connect_state()
        if state == 1:
            state_msg = "서버 연결 중"
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg)

    #잔고 실시간으로 갱신
    def timeout2(self):
        if self.checkBox.isChecked():
            self.check_balance()

    #현재가 실시간 갱신
    def timeout3(self):
        while self.checkBox_2.isChecked():
            self.present_price()
            
            
    #현재가격저장        
    def present_price(self):
        price = self.kiwoom.price
        self.lineEdit_3.setText(str(price))
     
    #주식 잔고 
    def check_balance(self):
        self.kiwoom.reset_opw00018_output()
        account_number = self.comboBox.currentText()

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

        while self.kiwoom.remained_data:
            time.sleep(0.2)
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 2, "2000")

        # opw00001
        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00001_req", "opw00001", 0, "2000")

        # balance
        item = QTableWidgetItem(self.kiwoom.d2_deposit)
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.tableWidget.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018_output['single'][i - 1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        self.tableWidget.resizeRowsToContents()

        # Item list
        item_count = len(self.kiwoom.opw00018_output['multi'])
        self.tableWidget_2.setRowCount(item_count)

        for j in range(item_count):
            row = self.kiwoom.opw00018_output['multi'][j]
            for i in range(len(row)):
                item = QTableWidgetItem(row[i])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(j, i, item)

        self.tableWidget_2.resizeRowsToContents()
    
    #선물 잔고
    def check_balance_2(self):
        self.kiwoom.reset_opw20006_output()
        account_number = self.comboBox.currentText()

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw20006_req", "opw20006", 0, "2000")

        while self.kiwoom.remained_data:
            time.sleep(0.2)
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw20006_req", "opw20006", 0, "2000")
            
        # opw00001
        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00001_req", "opw00001", 0, "2000")

        # balance
        item = QTableWidgetItem(self.kiwoom.d2_deposit)
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.tableWidget.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw20006_output['single'][i - 1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)

        self.tableWidget.resizeRowsToContents()

        # Item list
        item_count = len(self.kiwoom.opw20006_output['multi'])
        self.tableWidget_2.setRowCount(item_count)

        for j in range(item_count):
            row = self.kiwoom.opw20006_output['multi'][j]
            for i in range(len(row)):
                item = QTableWidgetItem(row[i])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(j, i, item)

        self.tableWidget_2.resizeRowsToContents()
        
        
    def trade_start(self):
        self.trade_set = True
        code = self.lineEdit.text()
        self.kiwoom.set_input_value("종목코드", code)
        self.kiwoom.comm_rq_data("opt50003_req", "opt50003", 0, "1000")

    def start_price_list(self):
        code = self.lineEdit.text()
        self.kiwoom.set_input_value("종목코드", code)
        self.kiwoom.comm_rq_data("opt50001_req", "opt50001", 0, "1000")
        
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
