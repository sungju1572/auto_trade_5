import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import pandas as pd
import sqlite3
import datetime
import numpy as np
from datetime import datetime

TR_REQ_TIME_INTERVAL = 0.2

class Kiwoom(QAxWidget):
    def __init__(self, ui):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()
        self.ui = ui
        self.state = "초기상태"
        self.pt = 0.5  #매도 매수 기준 point(0.5)
        self.reach_peak = 0 #2pt 찍었는지 구별 여부 (1이면 찍은것)
        self.sec_data = 0 #두번쨰 데이터
        self.trade_count = 0 #거래 횟수
        self.start_price = 0 #시가
        self.dic = {} #종목별로 정보 저장할 dict
        

        
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

    #실시간 타입 구독신청
    def SetRealReg(self, screen_no, code_list, fid_list, real_type):
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", 
                              screen_no, code_list, fid_list, real_type)
        

    def ready_trade(self, ticker, point, quantity):
        name = self.get_master_code_name(ticker)
        
        self.dic[name + '_name'] = name #종목명
        self.dic[name + '_ticker'] = ticker #종목 티커
        self.dic[name + '_point'] = float(point) #종목 포인트 기준 (몇 단위로 거래 들어갈건지)
        self.dic[name + '_quantity'] = int(quantity) #거래 수량
        self.dic[name + '_status'] = '초기상태' #종목별 현재상태
        self.dic[name + '_buy_count'] = 0 #거래 횟수 
        self.dic[name + '_refer'] = 0 #기준점
        self.dic[name + '_reach_peak'] = 0 #기준점에 도달했는지 여부 확인(0이면 미도달)
        self.dic[name + '_sec_data'] = 0 #기준점에 도달했을 때 가격
        self.dic[name + '_end_trade'] = 0 #거래 횟수 3번일때 거래 종료지점 체크
        
        
        
        self.ui.textEdit.append("거래준비완료 | 종목 : " + name)
        

####
    #실시간 조회관련 핸들
    def _handler_real_data(self, trcode, ret, data):
        
        #ui에서 계좌랑 종목코드 가져오기
        self.account = self.ui.comboBox.currentText()

        if ret == "선물시세":
            #체결시간
            self.time =  self.get_comm_real_data(trcode, 20)
            self.time = self.time[:2] + ":" + self.time[2:4] + ":" + self.time[4:6]
                      
    
            open_price = self.get_comm_real_data(trcode, 16) #시가
            price = self.get_comm_real_data(trcode, 10)       #현재가
            name = self.get_master_code_name(trcode)          #이름
            
          
            if price !="" and open_price != 0:
                price = float(price[1:])
                open_price = float(open_price[1:])
                
                
                self.dic[name + '_open_price'] = open_price  #dic에 시가 저장
                self.dic[name + '_price'] = price   #dic에 현재가 저장
                
               
                self.strategy(name)

            

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
        
        self.refer = self.start_price  #기준점 (시작은 시가로 시작)



        
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
    def strategy(self, name):
        try:
            list_1 = [k for k in self.dic.keys() if name in k ]
            
            name = self.dic[list_1[list_1.index(name + '_name')]]                #종목 이름
            ticker = self.dic[list_1[list_1.index(name + '_ticker')]]            #종목 티커
            point = self.dic[list_1[list_1.index(name + '_point')]]              #종목 포인트 기준
            quantity = self.dic[list_1[list_1.index(name + '_quantity')]]        #거래수량
            status = self.dic[list_1[list_1.index(name + '_status')]]            #현재 상태
            buy_count = self.dic[list_1[list_1.index(name + '_buy_count')]]      #거래 횟수
            open_price = self.dic[list_1[list_1.index(name + '_open_price')]]    #시가
            price = self.dic[list_1[list_1.index(name + '_price')]]               #현재가
            refer = self.dic[list_1[list_1.index(name + '_refer')]]               #기준점
            reach_peak = self.dic[list_1[list_1.index(name + '_reach_peak')]]     #기준점 도달 확인
            sec_data = self.dic[list_1[list_1.index(name + '_sec_data')]]         #기준점 도달 했을 때 가격
            end_trade = self.dic[list_1[list_1.index(name + '_end_trade')]]        #거래 종료
            
            self.ui.textEdit_2.append("종목 : " + str(name))  
            self.ui.textEdit_2.append("현재가 : " + str(price))   
            self.ui.textEdit_2.append("거래횟수 : " + str(buy_count))   
            self.ui.textEdit_2.append("기준점 : " + str(sec_data))   
            self.ui.textEdit_2.append("-----------------" )   
            

            
            #거래횟수 3번 이하일때 거래진행
            if buy_count <= 3 :
                
                #초기상태
                if status =="초기상태":
                    #long
                    if price > open_price + point:
                        self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "2", "3", quantity, "0", "")
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 롱진입 | 진입지점 : " + str(open_price + point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        self.dic[list_1[list_1.index(name + '_status')]]  = "롱포지션"
                        self.dic[list_1[list_1.index(name+'_refer')]] = open_price + point
                    #short    
                    elif price < open_price - point:
                        self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "1", "3", quantity, "0", "")
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 숏진입 | 진입지점 : " + str(open_price - point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        self.dic[list_1[list_1.index(name + '_status')]]  = "숏포지션"
                        self.dic[list_1[list_1.index(name+'_refer')]] = open_price - point
                        
                #매수상태
                elif status =="롱포지션":
                    #강제청산
                    if price < refer - 2*point:
                        self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "1", "3", quantity, "0", "")
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 롱청산 | 청산지점 : " + str(refer - 2*point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        self.dic[list_1[list_1.index(name+'_refer')]] = refer - 2*point
                        self.dic[list_1[list_1.index(name+'_buy_count')]] += 1 
                        self.dic[list_1[list_1.index(name+'_status')]] = "초기상태2"
                    
                    #2pt 도달 했는지 확인
                    elif price > refer + 3*point and reach_peak == 0 :
                        self.dic[list_1[list_1.index(name+'_sec_data')]] = refer + 4*point
                        self.dic[list_1[list_1.index(name+'_reach_peak')]] = 1
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 2pt 도달(long) | 도달지점 : " + str(refer + 3*point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        
                    if reach_peak == 1 :
                        #롱 익절
                        if price < sec_data - point :
                            self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "1", "3", quantity, "0", "")
                            self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 롱익절 | 익절지점 : " + str(sec_data - point))
                            self.ui.textEdit.append("현재가 : " + str(price))
                            self.ui.textEdit.append("                ")
                            self.dic[list_1[list_1.index(name+'_refer')]] = sec_data - point
                            
                            self.dic[list_1[list_1.index(name+'_reach_peak')]] = 0
                            self.dic[list_1[list_1.index(name+'_status')]] = "초기상태2"
                        
                        #0.5p 도달할때마다 기준점 갱신
                        elif price >= sec_data + point :
                            self.dic[list_1[list_1.index(name+'_sec_data')]] = sec_data + point
                            self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 롱 기준점 갱신 : " + str(sec_data + point))
                            self.ui.textEdit.append("현재가 : " + str(price))
                            self.ui.textEdit.append("                ")
                            

                
                #매도상태
                elif status =="숏포지션":
                    #강제청산
                    if price > refer + 2*point:
                        self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "2", "3", quantity, "0", "")
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 숏청산 | 청산지점 : " + str(refer + 2*point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        self.dic[list_1[list_1.index(name+'_refer')]] = refer + 2*point
                        self.dic[list_1[list_1.index(name+'_buy_count')]] += 1 
                        self.dic[list_1[list_1.index(name+'_status')]] = "초기상태2"
                    
                    #2pt 도달 했는지 확인
                    elif price < refer - 3*point and reach_peak == 0 :
                        self.dic[list_1[list_1.index(name+'_sec_data')]] = refer - 4*point
                        self.dic[list_1[list_1.index(name+'_reach_peak')]] = 1
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 2pt 도달(short) | 도달지점 : " + str(refer - 3*point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        
                        
                    if reach_peak == 1 :
                        #숏 익절
                        if price > sec_data + point :
                            self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "2", "3", quantity, "0", "")
                            self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 숏익절 | 익절지점 : " + str(sec_data + point))
                            self.ui.textEdit.append("현재가 : " + str(price))
                            self.ui.textEdit.append("                ")
                            self.dic[list_1[list_1.index(name+'_refer')]] = sec_data + point
                            
                            self.dic[list_1[list_1.index(name+'_reach_peak')]] = 0
                            self.dic[list_1[list_1.index(name+'_status')]] = "초기상태2"
                        
                        #0.5p 도달할때마다 기준점 갱신
                        elif price <= sec_data - point :
                            self.dic[list_1[list_1.index(name+'_sec_data')]] = sec_data + point
                            self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 숏 기준점 갱신 : " + str(sec_data - point))
                            self.ui.textEdit.append("현재가 : " + str(price))
                            self.ui.textEdit.append("                ")
                    

                        
                elif status == "초기상태2":
                    #long
                    if price > refer + point:
                        self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "2", "3", quantity, "0", "")
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 롱진입 | 진입지점 : " + str(refer + point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        self.dic[list_1[list_1.index(name + '_status')]]  = "롱포지션"
                        self.dic[list_1[list_1.index(name+'_refer')]] = refer + point
                    #short    
                    elif price < refer - point:
                        self.send_order_fo("send_order_fo_req", "0101", self.account, ticker, 1, "1", "3", quantity, "0", "")
                        self.ui.textEdit.append(str(self.time) + " | " + str(name) + " 숏진입 | 진입지점 : " + str(refer - point))
                        self.ui.textEdit.append("현재가 : " + str(price))
                        self.ui.textEdit.append("                ")
                        self.dic[list_1[list_1.index(name + '_status')]]  = "숏포지션"
                        self.dic[list_1[list_1.index(name+'_refer')]] = refer - point
                        
                    
            elif buy_count > 3 and end_trade == 0 :
                self.ui.textEdit.append("거래 종료 : " + str(name))
                self.dic[list_1[list_1.index(name+'_end_trade')]] = 1
                
                
                 
                
        

        
        except Exception as e:    # 예외 처리 (모든예외 일괄 처리)
            print('예외 발생 : ', e)
       
        

                

 

        



if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect() #연결
    

    
    