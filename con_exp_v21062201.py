import sys
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import Kiwoom_OpenAPI_Mod
from PyQt5 import uic
from PyQt5.QtGui import *
import should_have_def_2
import time
from datetime import datetime, timedelta
import telegram
import urllib.request
import requests
from bs4 import BeautifulSoup

bot_token = ""
solo_id = ""
red = "color: #FF0000;"
blue = "color: #0000FF;"
green = "color: #0FFF0F;"

con_exp = [["", "NONE"], ["", "NONE"], ["", "NONE"], ["", "NONE"], ["", "NONE"]]
con_exp_telegram = []
con_exp_flag = False
tr_busy_flag = False
form_class = uic.loadUiType("con_exp_gui.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        global bot_token
        global solo_id
        global con_exp

        super().__init__()
        self.setupUi(self)
        self.lb_login_msg.setText(f'로그인 중 / 계좌번호 : {account[0]}')

        KW.GetConditionLoad()
        con_exp_row = KW.GetConditionNameList()

        for i in range(len(con_exp_row)):
            con_exp[i] = con_exp_row[i]

        self.lb_con_exp_value.setText(str(len(con_exp_row)))

        self.lb_con_exp_name_1.setText(con_exp[0][1])
        self.lb_con_exp_name_2.setText(con_exp[1][1])
        self.lb_con_exp_name_3.setText(con_exp[2][1])
        self.lb_con_exp_name_4.setText(con_exp[3][1])
        self.lb_con_exp_name_5.setText(con_exp[4][1])

        # 메인
        self.le_telegram_bot_token.setText("1801941374:AAFWlXnwjnDs7YzfgW2VCbcu_yuu_2x8HZ8")
        self.le_telegram_id.setText("1822420422")

        bot_token = self.le_telegram_bot_token.text()
        solo_id = self.le_telegram_id.text()

        # 조건식 1
        self.con_exp_cycle_1 = 0
        self.le_con_exp_cycle_1.setText("1")
        self.tw_con_exp_1.setRowCount(0)  # 이거 안하면 rowCount로 못읽음

        self.con_exp_pixmap_1 = QPixmap()
        self.con_exp_timer_1 = QTimer()
        self.pb_con_exp_start_1.clicked.connect(self.con_exp_start_1)
        self.tw_con_exp_1.cellDoubleClicked.connect(self.con_exp_chart_view_1)

    def con_exp_chart_view_1(self):
        global tr_busy_flag

        row = self.tw_con_exp_1.currentRow()
        code = self.tw_con_exp_1.item(row, 1).text()

        self.con_exp_pixmap_1.load("chart/" + code + ".png")
        self.lb_con_exp_photo_1.setPixmap(self.con_exp_pixmap_1)

        self.lb_con_exp_stock_name_1.setText(stock_dic[code])

        tr_code_var = should_have_def_2.opt10001
        tr_code_var[1][0] = code

        while True:
            if not tr_busy_flag:
                data = tr_data_call("opt10001", tr_code_var)
                break

        print(data)

        self.le_con_exp_value0_1.setStyleSheet(red)
        self.le_con_exp_value0_1.setText("{:,}".format(int(sign_judge(data[0][0]))))

        self.le_con_exp_value1_1.setStyleSheet(blue)
        self.le_con_exp_value1_1.setText("{:,}".format(int(sign_judge(data[0][1]))))

        self.le_con_exp_value2_1.setText(digit(data[0][2]))

        self.le_con_exp_value3_1.setText(data[0][3])

        self.le_con_exp_value4_1.setText(data[0][4])

        self.le_con_exp_value5_1.setText(data[0][5])

        self.le_con_exp_value6_1.setText(data[0][6])

        self.le_con_exp_value7_1.setStyleSheet(color(data[0][7]))
        self.le_con_exp_value7_1.setText("{:,}".format(int(sign_judge(data[0][7]))))

        self.le_con_exp_value8_1.setStyleSheet(color(data[0][8]))
        self.le_con_exp_value8_1.setText("{:,}".format(int(sign_judge(data[0][8]))))

        self.le_con_exp_value9_1.setStyleSheet(color(data[0][9]))
        self.le_con_exp_value9_1.setText("{:,}".format(int(sign_judge(data[0][9]))))

        self.le_con_exp_value13_1.setText("{:,}".format(int(data[0][13])))

        self.lb_con_exp_now_1.setStyleSheet(color(data[0][10]))
        self.lb_con_exp_now_1.setText("{:,}".format(int(sign_judge(data[0][10]))))

        self.lb_con_exp_yest_1.setStyleSheet(color(data[0][11]))
        self.lb_con_exp_yest_1.setText("{:,}".format(int(sign_judge(data[0][11]))))

        self.lb_con_exp_rate_1.setStyleSheet(color(data[0][12]))

        if float(data[0][12]) > 0:
            self.lb_con_exp_rate_1.setText(data[0][12] + " ▲")

        elif float(data[0][12]) < 0:
            self.lb_con_exp_rate_1.setText(data[0][12] + " ▼")

        else:
            self.lb_con_exp_rate_1.setText(data[0][12] + " -")


    def con_exp_start_1(self):
        global con_exp_flag

        if self.con_exp_timer_1.isActive():
            self.con_exp_timer_1.stop()

        # 동일 조건식 1분에 1회 검색 가능
        else:
            self.con_exp_1()
            self.con_exp_timer_1.setInterval(int(self.le_con_exp_cycle_1.text()) * 61000)
            self.con_exp_timer_1.timeout.connect(self.con_exp_1)
            self.con_exp_timer_1.start()

    def con_exp_1(self):
        global con_exp_flag
        global con_exp_telegram

        con_exp_telegram_out = []

        if len(con_exp) >= 1:

            self.te_system_msg.append(f'[{time_now(3)}]  조건식 1 실행\n')

            while True:
                if not con_exp_flag:
                    con_code = con_exp_call(0)
                    print(con_code)
                    break

            self.cb_con_exp_chart_save_1.currentIndex()

            for _ in range(self.tw_con_exp_1.rowCount()):
                self.tw_con_exp_1.removeRow(0)

            self.tw_con_exp_1.setRowCount(len(con_code))

            for i in range(len(con_code)):
                self.tw_con_exp_1.setItem(i, 0, QTableWidgetItem((stock_dic[con_code[i]])))
                self.tw_con_exp_1.setItem(i, 1, QTableWidgetItem(con_code[i]))

                save_photo(con_code[i], self.cb_con_exp_chart_save_1.currentIndex())

            if self.cb_con_exp_telegram_1.isChecked():
                if self.cb_con_exp_overlap_1.isChecked():
                    con_exp_out = set_data_cal(1, con_code, con_exp_telegram)
                    con_exp_telegram = con_exp_telegram + con_exp_out

                    con_exp_telegram_out = telegram_link(con_exp_out)

                else:
                    con_exp_telegram_out = telegram_link(con_code)
                    con_exp_out = con_code

            if len(con_exp_telegram_out):
                if self.cb_con_exp_chart_1.isChecked():
                    for i in range(len(con_exp_telegram_out)):
                        telegram_text(f"{con_exp[0][1]} : {con_exp_telegram_out[i]}")

                        telegram_photo("chart/" + con_exp_out[i] + ".png")
                else:
                    telegram_msg(f"---- {con_exp[0][1]} ----", con_exp_telegram_out)




def color(data):
    if data.count('-'):
        return blue

    elif data.count('+'):
        return red

    else:
        return green


def digit(data):
    str_var = sign_judge(data)
    str_index_value = len(str_var)

    if str_index_value > 4:
        data = f"{str_var[:-4]}조{str_var[-4:]}억"

    else:
        data = str_var + "억"

    return data


def sign_judge(data):
    if data.count('-') or data.count('+'):
        data = (data[1:])  # 기호 삭제

    return data


def set_data_cal(cmd, data_1, data_2):
    s1 = set(data_1)
    s2 = set(data_2)

    if cmd == 1:  # 차집합
        return list(s1 - s2)


def con_exp_call(con_index):
    global con_exp_flag

    con_exp_flag = True

    if len(con_exp) >= con_index:
        while True:
            data = KW.SendCondition("0101", con_exp[con_index][1], con_exp[con_index][0], 0)
            if not data == 0:
                break

    con_exp_flag = False

    return data


def save_photo(code, i):
    if i == 0:
        url = "https://ssl.pstatic.net/imgfinance/chart/item/area/day/" + code + ".png"
        urllib.request.urlretrieve(url, "chart/" + code + ".png")

    elif i == 1:
        url = "https://ssl.pstatic.net/imgfinance/chart/item/area/week/" + code + ".png"
        urllib.request.urlretrieve(url, "chart/" + code + ".png")

    elif i == 2:
        url = "https://ssl.pstatic.net/imgfinance/chart/item/area/month3/" + code + ".png"
        urllib.request.urlretrieve(url, "chart/" + code + ".png")

    elif i == 3:
        url = "https://ssl.pstatic.net/imgfinance/chart/item/area/year/" + code + ".png"
        urllib.request.urlretrieve(url, "chart/" + code + ".png")


def sign_judge(data):
    if data.count('-') or data.count('+'):
        data = (data[1:])  # 기호 삭제

    return data


def etf_check(code):
    etf = 0

    for etf_name in should_have_def_2.etf:
        etf += code.count(etf_name)

    if etf:
        return 0

    else:
        return 1


def telegram_msg(str_1, str_2):
    telegram_msg_send = str_1
    telegram_msg_send += "\n"

    for name in str_2:
        telegram_msg_send += name
        telegram_msg_send += "\n"

    telegram_text(telegram_msg_send)


def telegram_text(msg):
    bot = telegram.Bot(token=bot_token)
    bot.sendMessage(chat_id=solo_id, text=msg, parse_mode="Markdown",
                    disable_web_page_preview=True)


def telegram_photo(path):
    bot = telegram.Bot(token=bot_token)
    bot.send_photo(solo_id, open(path, 'rb'))


def telegram_link(data):
    var = []
    for item in data:
        var.append("[%s](https://finance.naver.com/item/main.nhn?code=%s)"
                   % (stock_dic[item], item))
    return var


def telegram_file_send(file_path):
    bot = telegram.Bot(token=bot_token)
    bot.sendDocument(chat_id=solo_id, document=open(file_path, 'rb'))


def time_now(time_type):
    data = datetime.now()

    if time_type == 0:  # timestamp
        return data

    elif time_type == 1:  # 시:분
        return data.strftime("%H:%M")

    elif time_type == 2:
        return data.strftime('%Y%m%d')

    elif time_type == 3:
        return data.strftime("%H:%M:%S")

    else:
        return 0


def tr_data_call(tr_command, tr_name):
    global tr_busy_flag

    tr_busy_flag = True

    while True:
        KW.SetInputValue(tr_name[0], tr_name[1])

        data = KW.CommRqData(tr_command, tr_name[2])  # 거래량 증가 종목 요청

        if not data == 0:
            break

    tr_busy_flag = False

    return data


def stock_dic_update():
    kospi = KW.GetCodeListByMarket('0')
    kosdaq = KW.GetCodeListByMarket('10')

    dic_kospi = {kospi[0]: 'kospi_len'}
    dic_kosdaq = {kosdaq[0]: 'kosdaq_len'}

    for i in kospi:
        dic_kospi[i] = KW.GetMasterCodeName(i)

    for i in kosdaq:
        dic_kosdaq[i] = KW.GetMasterCodeName(i)

    dic_kospi.update(dic_kosdaq)

    return dic_kospi


if __name__ == "__main__":
    app = QApplication(sys.argv)
    KW = Kiwoom_OpenAPI_Mod.Kiwoom_Quant()

    KW.Login(block=True)

    print("로그인 완료")

    stock_dic = stock_dic_update()

    stock_code_name = list(stock_dic.keys())
    stock_name = list(stock_dic.values())

    account = KW.GetLoginInfo("ACCNO")

    KW.GetConditionLoad()

    mywindow = MyWindow()
    mywindow.show()
    app.exec_()

# 공시 가져오는거???????????