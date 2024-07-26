import sys
import time

import should_have_def
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import pythoncom  # COM 관련
import threading


class Kiwoom_Quant:
    def __init__(self):
        self.KW = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")  # ProgID를 클래스 생성자로 전달하여 인스턴스 생성
        #        self.KQ = QAxWidget("A1574A0D-6BFA-4BD7-9020-DED88711818D")    # 요렇게 해도 됨
        self.connected = False  # 연결 상태 확인 용 변수
        self.tr_ok = False
        self.trcode = ""
        self.rqname = ""
        self.tr_data = []
        self.screen = "0101"
        self.output_name = []
        self.tr_command_flag = True
        self.condition_loaded = False
        self.tr_condition_loaded = False
        self.Event_init()  # 수신 이벤트 init

    def Event_init(self):
        self.KW.OnEventConnect.connect(self._OnEventConnect)  # 서버 접속 관련 이벤트 발생 시 해당 함수로 이동
        self.KW.OnReceiveTrData.connect(self._OnReceiveTrData)
        self.KW.OnReceiveMsg.connect(self._OnReceiveMsg)
        self.KW.OnReceiveChejanData.connect(self._OnReceiveChejanData)
        self.KW.OnReceiveConditionVer.connect(self._handler_condition_load)
        self.KW.OnReceiveTrCondition.connect(self._handler_tr_condition)

    def _OnReceiveMsg(self, screen, rqname, trcode, msg):
        print("Msg : " + msg)
        print("")

    def _OnEventConnect(self, err_code):  # err_code 0이면 정상 처리
        if err_code == 0:
            self.connected = True  # login 성공

    def _OnReceiveTrData(self, screen, rqname, trcode, record, next):
        if self.tr_command_flag:
            rows = self.KW.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)

            if rows == 0:
                rows = 1

            data_list = []

            for row in range(rows):
                row_data = []

                for item in self.output_name:
                    row_data.append(
                        self.KW.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, row,
                                            item))

                for i in range(len(row_data)):
                    str_val = row_data[i]
                    str_val = str_val.strip()
                    row_data[i] = str_val

                data_list.append(row_data)

            self.tr_data = data_list
        else:
            self.tr_data = []
            self.tr_data.append(screen)
            self.tr_data.append(rqname)
            self.tr_data.append(trcode)
            self.tr_data.append(record)
            self.tr_data.append(next)

        self.tr_ok = True

    def _OnReceiveChejanData(self, gubun, cnt, fid):
        # self.KW.dynamicCall("GetCommData(int)", 9001)
        # self.tr_data.append(self.KW.dynamicCall("GetCommData(int)", 9001))
        # self.tr_data.append(self.KW.dynamicCall("GetCommData(int)", 910))
        # self.tr_data.append(self.KW.dynamicCall("GetCommData(int)", 911))
        # print("Chejan Data = %s, %d, %s" % (gubun, cnt, fid))
        pass

###########
    def _handler_condition_load(self, ret, msg):
        if ret == 1:
            self.condition_loaded = True

    def _handler_tr_condition(self, screen_no, code_list, cond_name, cond_index, next):
        codes = code_list.split(';')[:-1]
        self.tr_condition_data = codes
        self.tr_condition_loaded = True
############
    # -------------------------------------------------------------------------
    # OpenAPI+
    # -------------------------------------------------------------------------

    #   ---- 로그인 ----
    def Login(self, block=True):  # 로그인
        self.KW.dynamicCall("CommConnect()")  # OCX 방식은 dynamicCall을 사용하여 메서드 호출

        if block:
            while not self.connected:  # connected가 False면 not이므로 while 실행 / Event 후 True로 변경됨
                pythoncom.PumpWaitingMessages()  # COM에 연결된 객체 메세지 출력 <Class 'int'>

    #   ---- 로그 아웃 ----
    def Logout(self):
        self.KW.dynamicCall("void CommTerminate()")

    #   ---- 로그인 정보 요청 ----
    def GetLoginInfo(self, tag):  # 로그인 정보
        data = self.KW.dynamicCall("GetLoginInfo(QString)", tag)  # BSTR인자는 QString으로 보내야함
        # dynamicCall("문자열", 인자1, 인자2, ...)
        if tag == "ACCNO":
            return data.split(';')[:-1]
        else:
            return data

    #   --- 종목 코드 반환 ----
    def GetCodeListByMarket(self, market):
        data = self.KW.dynamicCall("GetCodeListByMarket(QString)", market)
        tokens = data.split(';')[:-1]
        return tokens

    #   ---- 종목 이름 반환 ----
    def GetMasterCodeName(self, code):
        data = self.KW.dynamicCall("GetMasterCodeName(QString)", code)
        return data

    #   ---- TR 요청 함수 Input ----
    def SetInputValue(self, value, code):
        for i in range(len(value)):
            self.KW.dynamicCall("SetInputValue(QString, QString)", value[i], code[i])

    def CommRqData(self, trcode, output_list):
        self.tr_command_flag = True

        self.tr_ok = False

        self.output_name = output_list

        if self.screen == "0101":
            self.screen = "0102"

        else:
            self.screen = "0101"

        self.KW.dynamicCall("CommRqData(QString, QString, int, QString)", trcode + "_req", trcode, 0, self.screen)

        t_end = time.time() + 5  # 5초 타이머 선언

        while time.time() < t_end:  # 5초
            pythoncom.PumpWaitingMessages()  # pass 적으면 못빠져나오더라

            if self.tr_ok:
                break

        time.sleep(0.3)

        self.output_name = []

        if self.tr_ok:
            return self.tr_data

        else:
            return 0

    def SendOrder(self, order_type, code, qty):
        self.tr_command_flag = False

        self.tr_ok = False

        self.KW.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int ,QString, QString)",
                            ["send_order_req", "0101", "8165838111", order_type, code, qty, 0, "03", ""])

        t_end = time.time() + 5  # 5초 타이머 선언

        while time.time() < t_end:  # 3초
            pythoncom.PumpWaitingMessages()  # pass 적으면 못빠져나오더라

            if self.tr_ok:
                break

        time.sleep(0.3)

        return self.tr_data
#################
    def GetConditionLoad(self, block=True):
        self.condition_loaded = False
        self.KW.dynamicCall("GetConditionLoad()")
        if block:
            while not self.condition_loaded:
                pythoncom.PumpWaitingMessages()

    def GetConditionNameList(self):
        data = self.KW.dynamicCall("GetConditionNameList()")
        conditions = data.split(";")[:-1]

        # [('000', 'perpbr'), ('001', 'macd'), ...]
        result = []
        for condition in conditions:
            cond_index, cond_name = condition.split('^')
            result.append((cond_index, cond_name))

        return result

    def SendCondition(self, screen, cond_name, cond_index, search):
        self.tr_condition_loaded = False
        self.KW.dynamicCall("SendCondition(QString, QString, int, int)", screen, cond_name, cond_index, search)

        while not self.tr_condition_loaded:
            pythoncom.PumpWaitingMessages()

        return self.tr_condition_data
############################
if __name__ == "__main__":
    app = QApplication(sys.argv)  # 객체를 실행하기 위한 절대경로(sys.argv) 지정

    login_test = Kiwoom_Quant()
    login_test.Login(block=True)
    print("로그인 완료")

    print(login_test.GetLoginInfo("USER_ID"))

    app.exec_()
