import sys

from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import Kiwoom_OpenAPI_Mod
from PyQt5 import uic, QtGui
import should_have_def_2
import time
from datetime import datetime, timedelta
import telegram
import urllib.request
import requests
from bs4 import BeautifulSoup

tr_busy_flag = False
account = ""
telegram_high_stock = []
telegram_real_time_stock = []

form_class = uic.loadUiType("should_have_gui.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.lb_login_msg.setText(f'로그인 중 / 계좌번호 : {account[0]}')

        # 거래량 / 전고점 돌파 Init
        self.top_trade_stock_cycle_value = 0
        self.le_top_trade_stock_cycle.setText("60")
        self.tw_top_trade_stock.setRowCount(0)  # 이거 안하면 rowCount로 못읽음

        self.top_trade_stock_timer = QTimer()
        self.pb_top_trade_stock.clicked.connect(self.top_trade_stock_timer_start)
        self.le_top_trade_stock_cycle.textChanged.connect(self.top_trade_stock_cycle_changed)

        self.top_trade_stock_cycle_changed()

        # 실시간 상승 Init
        self.real_time_stock_cycle_value = 0
        self.real_time_stock_volume_value = 0
        self.real_time_stock_rate_value = 0
        self.le_real_time_stock_volume.setText("70")
        self.le_real_time_stock_rate.setText("1")
        self.tw_real_time_stock.setRowCount(0)

        self.real_time_stock_timer = QTimer()
        self.pb_real_time_stock.clicked.connect(self.real_time_stock_timer_start)
        self.cb_real_time_stock_cycle.currentIndexChanged.connect(self.real_time_stock_cycle_changed)
        self.le_real_time_stock_volume.textChanged.connect(self.real_time_stock_volume_changed)
        self.le_real_time_stock_rate.textChanged.connect(self.real_time_stock_rate_changed)

        self.real_time_stock_cycle_changed()
        self.real_time_stock_volume_changed()
        self.real_time_stock_rate_changed()

        # 실시간 주가 감시 Init
        self.real_time_cost_rate_value_1 = 0
        self.real_time_cost_rate_value_2 = 0
        self.le_real_time_cost_rate_1.setText("1")
        self.le_real_time_cost_rate_2.setText("1")
        self.tw_real_time_stock_low.setRowCount(0)

        self.real_time_cost_timer = QTimer()
        self.pb_real_time_cost.clicked.connect(self.real_time_cost_timer_start)
        self.le_real_time_cost_rate_1.textChanged.connect(self.real_time_cost_rate_changed_1)
        self.le_real_time_cost_rate_2.textChanged.connect(self.real_time_cost_rate_changed_2)

        self.real_time_cost_rate_changed_1()
        self.real_time_cost_rate_changed_2()

    # 실시간 주가 감시 함수
    def real_time_cost_timer_start(self):
        if self.real_time_cost_timer.isActive():
            self.real_time_cost_timer.stop()

        else:
            self.real_time_cost()
            self.real_time_cost_timer.setInterval(10000)
            self.real_time_cost_timer.timeout.connect(self.real_time_cost)
            self.real_time_cost_timer.start()

    def real_time_cost_rate_changed_1(self):
        self.real_time_cost_rate_value_1 = self.le_real_time_cost_rate_1.text()

    def real_time_cost_rate_changed_2(self):
        self.real_time_cost_rate_value_2 = self.le_real_time_cost_rate_2.text()

    def real_time_cost(self):
        row_num = self.tw_real_time_stock.rowCount()
        delete_code = []

        self.te_system_msg.append(f'[{time_now(3)}]  실시간 주가 검사 시작\n')

        for i in range(row_num):
            print(f'real_time_cost / for i in range(row_num): {i}')
            table_code = self.tw_real_time_stock.item(i, 6)
            code = table_code.text()
            table_cost = self.tw_real_time_stock.item(i, 3)
            before_cost = float(table_cost.text())
            table_high_rate = self.tw_real_time_stock.item(i, 5)
            high_rate = float(table_high_rate.text())

            real_cost = int(get_price(code))
            print(f'real_time_cost / for i in range(row_num): / real_cost {real_cost}')

            rate = round((float(real_cost - before_cost) * 100) / before_cost, 2)

            self.tw_real_time_stock.setItem(i, 4, QTableWidgetItem(str(real_cost)))
            self.tw_real_time_stock.setItem(i, 2, QTableWidgetItem(str(rate)))

            if rate > 0:
                self.tw_real_time_stock.item(i, 2).setForeground(QBrush(QColor(255, 0, 0)))

            elif rate < 0:
                self.tw_real_time_stock.item(i, 2).setForeground(QBrush(QColor(0, 0, 255)))

            else:
                self.tw_real_time_stock.item(i, 2).setForeground(QBrush(QColor(0, 255, 0)))

            if high_rate <= rate:
                self.tw_real_time_stock.setItem(i, 5, QTableWidgetItem(str(rate)))

            if self.cb_high_cost_1.isChecked():
                if (high_rate - float(self.real_time_cost_rate_value_1)) > rate:
                    delete_code.append(i)

            elif self.cb_high_cost_2.isChecked():
                if rate < (float(self.real_time_cost_rate_value_2)) * -1:
                    delete_code.append(i)

        if not len(delete_code) == 0:
            print(f'real_time_cost / if not len(delete_code) == 0: {len(delete_code)}')
            row_num = self.tw_real_time_stock_low.rowCount()

            self.tw_real_time_stock_low.setRowCount(row_num + len(delete_code))

            # if self.cb_real_time_stock_telegram.isChecked():
            #     telegram_text("---- 실시간 하락 전환 종목 ----")

            for i in range(len(delete_code)):
                print(f'real_time_cost / for i in range(len(delete_code)): {i}')
                self.tw_real_time_stock_low.setItem(row_num + i, 0, QTableWidgetItem(time_now(1)))
                self.tw_real_time_stock_low.setItem(row_num + i, 1,
                                                    QTableWidgetItem((self.tw_real_time_stock.item(delete_code[i] - i, 1)).text()))
                self.tw_real_time_stock_low.setItem(row_num + i, 2,
                                                    QTableWidgetItem((self.tw_real_time_stock.item(delete_code[i] - i, 2)).text()))
                self.tw_real_time_stock_low.setItem(row_num + i, 3,
                                                    QTableWidgetItem((self.tw_real_time_stock.item(delete_code[i] - i, 3)).text()))
                self.tw_real_time_stock_low.setItem(row_num + i, 4,
                                                    QTableWidgetItem((self.tw_real_time_stock.item(delete_code[i] - i, 4)).text()))
                self.tw_real_time_stock_low.setItem(row_num + i, 5,
                                                    QTableWidgetItem((self.tw_real_time_stock.item(delete_code[i] - i, 5)).text()))
                self.tw_real_time_stock_low.setItem(row_num + i, 6,
                                                    QTableWidgetItem((self.tw_real_time_stock.item(delete_code[i] - i, 6)).text()))

                code = (self.tw_real_time_stock.item(delete_code[i] - i, 6)).text()

                if self.cb_real_time_stock_telegram.isChecked():
                    save_photo(code)
                    telegram_text("[%s](https://finance.naver.com/item/main.nhn?code=%s) ▼"
                                  % (stock_dic[code], code))

                    telegram_photo("chart/" + code + ".png")

                self.tw_real_time_stock.removeRow(delete_code[i] - i)

        self.te_system_msg.append('')

    # 실시간 상승 함수
    def real_time_stock_timer_start(self):
        if self.real_time_stock_timer.isActive():
            self.real_time_stock_timer.stop()

        else:
            self.real_time_stock()
            self.real_time_stock_timer.setInterval(int(self.real_time_stock_cycle_value) * 60000)
            self.real_time_stock_timer.timeout.connect(self.real_time_stock)
            self.real_time_stock_timer.start()

    def real_time_stock_cycle_changed(self):
        self.real_time_stock_cycle_value = self.cb_real_time_stock_cycle.currentText()

    def real_time_stock_volume_changed(self):
        self.real_time_stock_volume_value = self.le_real_time_stock_volume.text()

    def real_time_stock_rate_changed(self):
        self.real_time_stock_rate_value = self.le_real_time_stock_rate.text()

    def real_time_stock(self):
        global tr_busy_flag
        global telegram_real_time_stock

        real_time_stock_value_now = []
        table_list_code = []

        self.te_system_msg.append(f'[{time_now(3)}]  실시간 상승 종목 찾기 시작\n')

        list_var_1 = should_have_def_2.opt10023
        list_var_1[1][4] = self.real_time_stock_cycle_value

        while True:
            if not tr_busy_flag:
                data = tr_data_call("opt10023", list_var_1)
                break

        row_num = self.tw_real_time_stock.rowCount()

        for i in range(row_num):
            print(f'real_time_stock / for i in range(row_num): {i}')
            table_code = self.tw_real_time_stock.item(i, 6)
            table_list_code.append(table_code.text())

        for i in range(len(data)):
            if etf_check(data[i][1]):
                if float(data[i][8]) > float(self.real_time_stock_volume_value):
                    if data[i][3] == "2":
                        if table_list_code.count(data[i][0]) == 0:
                            self.te_system_msg.append(f'[{data[i][1]}]  {data[i][2]} {data[i][4]} {data[i][8]}')

                            list_var_1 = should_have_def_2.opt10080
                            list_var_1[1][0] = data[i][0]
                            list_var_1[1][1] = self.real_time_stock_cycle_value
                            min_candle = 0

                            while True:
                                if not tr_busy_flag:
                                    min_candle_raw = tr_data_call("opt10080", list_var_1)
                                    break

                            for j in range(5):
                                print(f'real_time_stock / for i in range(len(data)) / for j in range(5): {j}')
                                min_candle += int(sign_judge(min_candle_raw[j + 1][0]))

                            min_candle = min_candle // 5

                            before_cost = min_candle + ((min_candle // 100) * int(self.real_time_stock_rate_value))

                            if int(sign_judge(data[i][2])) > before_cost:
                                real_time_stock_value_now.append(data[i])
                                save_photo(data[i][0])

        self.tw_real_time_stock.setRowCount(row_num + len(real_time_stock_value_now))

        for i in range(len(real_time_stock_value_now)):
            print(f'real_time_stock / for i in range(len(real_time_stock_value_now)): {i}')
            self.tw_real_time_stock.setItem(row_num + i, 0, QTableWidgetItem(time_now(1)))
            self.tw_real_time_stock.setItem(row_num + i, 1, QTableWidgetItem(real_time_stock_value_now[i][1]))
            # self.tw_real_time_stock.setItem(row_num + i, 2, QTableWidgetItem(real_time_stock_value_now[i][]))
            self.tw_real_time_stock.setItem(row_num + i, 3, QTableWidgetItem(sign_judge(real_time_stock_value_now[i][2])))
            # self.tw_real_time_stock.setItem(row_num + i, 5, QTableWidgetItem("-"))
            self.tw_real_time_stock.setItem(row_num + i, 5, QTableWidgetItem("0.0"))
            self.tw_real_time_stock.setItem(row_num + i, 6, QTableWidgetItem(real_time_stock_value_now[i][0]))

        if self.cb_real_time_stock_telegram.isChecked():
            table_list_code = []
            row_num = self.tw_real_time_stock.rowCount()

            for i in range(row_num):
                table_code = self.tw_real_time_stock.item(i, 6)
                table_list_code.append(table_code.text())

            real_time_stock_var = set_data_cal(1, table_list_code, telegram_real_time_stock)
            telegram_real_time_stock += real_time_stock_var

            if not len(real_time_stock_var) == 0:
                # telegram_text("---- 실시간 상승 종목 ----")
                for item in real_time_stock_var:
                    telegram_text("[%s](https://finance.naver.com/item/main.nhn?code=%s) ▲"
                                  % (stock_dic[item], item))

                    telegram_photo("chart/" + item + ".png")

        self.te_system_msg.append("")

    # 거래량 / 전고점 돌파 함수
    def top_trade_stock_timer_start(self):
        if self.top_trade_stock_timer.isActive():
            self.top_trade_stock_timer.stop()

        else:
            self.top_trade_stock()
            self.top_trade_stock_timer.setInterval(int(self.top_trade_stock_cycle_value) * 1000)
            self.top_trade_stock_timer.timeout.connect(self.top_trade_stock)
            self.top_trade_stock_timer.start()

    def top_trade_stock_cycle_changed(self):
        self.top_trade_stock_cycle_value = self.le_top_trade_stock_cycle.text()

    def top_trade_stock(self):
        global tr_busy_flag
        global telegram_high_stock

        high_stock_code = []
        high_stock_name = []
        high_stock_value = []
        high_stock_value_rate = []
        table_list_code = []

        self.te_system_msg.append(f'[{time_now(3)}]  거래량 / 전고점 돌파 종목 찾기 시작\n')

        while True:
            if not tr_busy_flag:
                high_value = tr_data_call("opt10016", should_have_def_2.opt10016)
                break

        while True:
            if not tr_busy_flag:
                high_volume = tr_data_call("opt10024", should_have_def_2.opt10024)
                break

        row_num = self.tw_top_trade_stock.rowCount()

        for i in range(row_num):
            print(f'top_trade_stock / for i in range(row_num): {i}')
            data = self.tw_top_trade_stock.item(i, 4)
            table_list_code.append(data.text())

        for item_value in high_value:
            for item_volume in high_volume:
                if item_value[0] == item_volume[0] and etf_check(item_value[1]):
                    if table_list_code.count(item_value[0]) == 0:
                        print(f'top_trade_stock / for item_value in high_value: {item_value[0]}')
                        high_stock_code.append(item_value[0])
                        high_stock_name.append(item_value[1])
                        high_stock_value_rate.append(item_value[2])
                        high_stock_value.append(item_value[3])

        self.tw_top_trade_stock.setRowCount(row_num + len(high_stock_code))

        for i in range(len(high_stock_code)):
            print(f'top_trade_stock /for i in range(len(high_stock_code)): {i}')
            self.tw_top_trade_stock.setItem(row_num + i, 0, QTableWidgetItem(time_now(1)))
            self.tw_top_trade_stock.setItem(row_num + i, 1, QTableWidgetItem(high_stock_name[i]))
            self.tw_top_trade_stock.setItem(row_num + i, 2, QTableWidgetItem(high_stock_value[i]))
            self.tw_top_trade_stock.setItem(row_num + i, 3, QTableWidgetItem(high_stock_value_rate[i]))
            self.tw_top_trade_stock.setItem(row_num + i, 4, QTableWidgetItem(high_stock_code[i]))

        if self.cb_top_trade_stock_telegram.isChecked():
            high_stock_var = set_data_cal(1, table_list_code, telegram_high_stock)
            telegram_high_stock = telegram_high_stock + high_stock_var

            high_stock_var = telegram_link(high_stock_var)

            if not len(high_stock_var) == 0:
                telegram_msg("---- 전일 거래량 100% 및 52주 최고가 돌파 종목 ----", high_stock_var)


def get_price(code):
    url = "https://finance.naver.com/item/main.nhn?code=" + code
    result = requests.get(url)
    bs_obj = BeautifulSoup(result.content, "html.parser")
    no_today = bs_obj.find("p", {"class": "no_today"})
    blind_now = no_today.find("span", {"class": "blind"})

    return blind_now.text.replace(',', '')


def save_photo(code):
    url = "https://ssl.pstatic.net/imgfinance/chart/item/area/day/" + code + ".png"
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
    telegram_msg_send = ""
    # telegram_msg_send += "\n"

    for name in str_2:
        telegram_msg_send += name
        telegram_msg_send += "\n"

    telegram_text(telegram_msg_send)


def telegram_text(msg):
    bot = telegram.Bot(token=should_have_def_2.bot_token)
    bot.sendMessage(chat_id=should_have_def_2.solo_id, text=msg, parse_mode="Markdown",
                    disable_web_page_preview=True)


def telegram_photo(path):
    bot = telegram.Bot(token=should_have_def_2.bot_token)
    bot.send_photo(should_have_def_2.solo_id, open(path, 'rb'))


def telegram_link(data):
    var = []
    for item in data:
        var.append("[%s](https://finance.naver.com/item/main.nhn?code=%s)"
                   % (stock_dic[item], item))
    return var


def telegram_file_send(file_path):
    bot = telegram.Bot(token=should_have_def_2.bot_token)
    bot.sendDocument(chat_id=should_have_def_2.solo_id, document=open(file_path, 'rb'))


def set_data_cal(cmd, data_1, data_2):
    s1 = set(data_1)
    s2 = set(data_2)

    if cmd == 1:  # 차집합
        return list(s1 - s2)


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

    conditions = KW.GetConditionNameList()
    print(conditions)

    # 0번 조건식에 해당하는 종목 리스트 출력
    condition_index = conditions[0][0]
    condition_name = conditions[0][1]
    codes = KW.SendCondition("0101", condition_name, condition_index, 0)

    print(codes)



    mywindow = MyWindow()
    mywindow.show()
    app.exec_()
