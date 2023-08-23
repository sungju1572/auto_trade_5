import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from Kiwoom import *
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

        self.lineEdit_6.textChanged.connect(self.code_changed)
        self.lineEdit_9.textChanged.connect(self.code_changed_2)

        self.pushButton_4.clicked.connect(self.check_balance_2)
        self.pushButton_7.clicked.connect(self.trade_start)
        
        
        self.gudoc_status = 0 #구독 상태 (0이면 구독된 종목개수 0개)

        

        
        

    def code_changed(self):
        code = self.lineEdit_6.text()
        name = self.kiwoom.get_master_code_name(code)
        self.lineEdit_2.setText(name)
        
        
    def code_changed_2(self):
        code = self.lineEdit_9.text()
        name = self.kiwoom.get_master_code_name(code)
        self.lineEdit_5.setText(name) 

    
    #계좌설정
    def set_account(self):
        account = self.comboBox.currentText()
        return account
        



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
            self.check_balance_2()
    

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
        
    #구독 개수 상태에따라 종목 구독
    def set_gudoc(self, code):
        if self.gudoc_status == 0 :
            self.kiwoom.SetRealReg(1000 , code, "20;10", "0")
            self.gudoc_status = 1
        elif self.gudoc_status == 1 :
            self.kiwoom.SetRealReg(1000 , code, "20;10", "1")
        
        self.textEdit.append("구독성공 : " + code)
        
        
    
        
    def trade_start(self):
        print("--대기중--")
        
        #코스피 200 선물 구독
        if self.checkBox_2.isChecked():
            point = self.lineEdit_3.text()
            code = self.lineEdit_6.text()
            quantity = self.lineEdit_7.text()
            self.kiwoom.ready_trade(code, point, quantity)
            self.set_gudoc(code)

        
        if self.checkBox_3.isChecked():
            point = self.lineEdit_4.text()
            code = self.lineEdit_9.text()
            quantity = self.lineEdit_8.text()
            self.kiwoom.ready_trade(code, point, quantity)
            self.set_gudoc(code)
            
        
    
        
        
        #self.kiwoom.set_input_value("종목코드", code)
        #self.kiwoom.comm_rq_data("opt50003_req", "opt50003", 0, "1000")


        
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()