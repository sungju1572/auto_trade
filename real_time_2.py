import sys 
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import datetime


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real")
        self.setGeometry(300, 300, 300, 400)

        btn = QPushButton("Register", self)
        btn.move(20, 20)
        btn.clicked.connect(self.btn_clicked)

        btn2 = QPushButton("DisConnect", self)
        btn2.move(20, 100)
        btn2.clicked.connect(self.btn2_clicked)

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.ocx.OnEventConnect.connect(self._handler_login)
        self.ocx.OnReceiveRealData.connect(self._handler_real_data)
        self.CommmConnect()

    def btn_clicked(self):
        self.ocx.dynamicCall("SetInputValue(QString, QString)", "종목코드", "101R9000")
        self.ocx.dynamicCall("CommRqData(QString, QString, int, QString)", "opt50001_req", "opt50001", 0, "1000")
        
        
    def btn2_clicked(self):
        self.DisConnectRealData("1000")

    def CommmConnect(self):
        self.ocx.dynamicCall("CommConnect()")
        self.statusBar().showMessage("login 중 ...")

    def _handler_login(self, err_code):
        if err_code == 0:
            self.statusBar().showMessage("login 완료")


    def _handler_real_data(self, code, real_type, data):
 #       print(code, real_type, data)
            # 체결 시간 
        time =  self.GetCommRealData(code, 20)
        date = datetime.datetime.now().strftime("%Y-%m-%d ")

        if time != "":
            time =  datetime.datetime.strptime(date + time, "%Y-%m-%d %H%M%S")
        print(time, end=" ")

            # 현재가 
        self.price =  self.GetCommRealData(code, 10)
        print(self.price)


    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.ocx.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)

    def DisConnectRealData(self, screen_no):
        self.ocx.dynamicCall("DisConnectRealData(QString)", screen_no)

    def GetCommRealData(self, code, fid):
        data = self.ocx.dynamicCall("GetCommRealData(QString, int)", code, fid) 
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()