import time
import pywhatkit
import webbrowser
# import pyautogui
from pynput.keyboard import Key, Controller
from tkinter import messagebox
from WhatsappSendMessage import get_date_of_week

keyboard = Controller()

import pandas as pd




def send_whatsapp_message(msg: str, phone_numbers: list):
    try:
        for phone_no in phone_numbers:
            print(type(phone_no))
            pywhatkit.sendwhatmsg_instantly(
                phone_no=phone_no,
                message=msg,
                tab_close=True
            )

            # time.sleep(1)
            # pyautogui.click()
            # time.sleep(1)
            # keyboard.press(Key.enter)
            # keyboard.release(Key.enter)
            print(f"Message sent to {phone_no}!")

    except Exception as e:
        print(str(e))


class WebWhatsappSendMessage:
    staying_leaving_columns = ['寝室', '姓名', '日期', '时间', '监护人电话', '离舍', '其他信息']

    def send(self, phone_no: str, msg: str):
        print(type(phone_no), phone_no, type(msg), msg)
        try:
            pywhatkit.sendwhatmsg_instantly(
                phone_no=phone_no,
                message=msg,
                wait_time=self.delay,
                tab_close=True
            )
        except Exception as e:
            print(str(e))

    def read_excel_file(self):
        data_frame = pd.read_excel(self.source_file, index_col='姓名', dtype={'监护人电话': str})
        return data_frame

    def set_value(self, txt_source_file, float_delay=15):
        self.source_file = txt_source_file
        self.delay = float_delay

    def set_stop(self, status: bool):
        self.stop_sending = status

    def __init__(self):
        self.source_file = ""
        self.stop_sending = False
        self.delay = 13

    def run(self):
        self.set_stop(False)
        webbrowser.open('https://www.google.com')
        df = self.read_excel_file()

        for name, rows in df.iterrows():
            message = ''
            if pd.isna(rows['监护人电话']):
                continue

            if not pd.isna(rows['其他信息']):
                message = str(rows['其他信息'])
                self.send(rows['监护人电话'], message)
                continue

            if rows['离舍'] == 0:
                message = f"{name} 同学 {rows['寝室']} 于 {str(rows['日期'])[:10]} 本周留舍 (留校)。\n敬请家长/监护人关注。"
                self.send(rows['监护人电话'], message)

            elif rows['离舍'] == 1:
                text_date = ' (星期' + get_date_of_week(rows['日期']) + ')'
                timing = str(rows['日期'])[:10] + text_date + ' ' + str(rows['时间'])[:5]
                message = f'{name} 同学 {rows['寝室']} 于 {timing} 离校 (回家)。\n敬请家长/监护人关注。'
                self.send(rows['监护人电话'], message)

            if self.stop_sending:
                messagebox.showinfo(message=f'已停止发送。\n最后一位信息接收者：{name}')
                break


if __name__ == "__main__":
    # List of phone numbers of your friends
    # friend_numbers = ["+6586518193", "+6586518193", "+6586518193", ]

    #message = "Hey bro, have you completed your Python task?"

    #send_whatsapp_message(msg=message, phone_numbers=friend_numbers)
    a = WebWhatsappSendMessage()
    a.set_value('w.xlsx')
    a.run()
    #a = WebWhatsappSendMessage()
    #a.set_value('w.xlsx')
    #df = a.read_excel_file()
    #print(df.head())
