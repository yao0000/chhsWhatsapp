import threading
import time

import pyautogui
import pywhatkit
import webbrowser
import pandas as pd
from pynput.keyboard import Controller
from tkinter import messagebox

import Log
from WhatsappSendMessage import get_date_of_week

from MaterialSource import (
    user_not_found_images_paths
)

keyboard = Controller()

delay = 8


def detect_error(phone_number: str):
    print('enter detect error function')
    accuracy = 0.8
    time.sleep(delay)

    for image in user_not_found_images_paths:
        try:
            coords = pyautogui.locateOnScreen(image, confidence=accuracy)

            if coords is not None:
                Log.log_message(_time=time.localtime(), receiver=phone_number, message='用户不存在')
                print(f'{phone_number} 不存在')
                break
        except Exception as err:
            print(f"err; {str(err)}")


def whatsapp(phone_no: str, msg: str):
    pywhatkit.sendwhatmsg_instantly(phone_no=phone_no, message=msg, tab_close=True, wait_time=delay)


def send_whatsapp_message(msg: str, phone_numbers: list):
    try:
        for phone_no in phone_numbers:
            whatsapp_thread = threading.Thread(target=whatsapp, args=(phone_no, msg))
            whatsapp_thread.start()

            detect_error(phone_no)
            whatsapp_thread.join()

            print(f"Message sent to {phone_no}!")

    except Exception as err:
        print(str(err))


class WebWhatsappSendMessage:
    def invalid_account_check(self, phone_number: str):
        accuracy = 0.8
        time.sleep(self.delay)

        for image in user_not_found_images_paths:
            try:
                coords = pyautogui.locateOnScreen(image, confidence=accuracy)

                if coords is not None:
                    Log.log_message(_time=time.localtime(),
                                    receiver=phone_number,
                                    message='检测到该用户不存在。')
                    print(f'{phone_number} 用户不存在')
                    return
            except Exception as err:
                print(f'error: {str(err)}')

    def whatsapp_send(self, phone_number: str, msg: str):
        pywhatkit.sendwhatmsg_instantly(phone_no=phone_number,
                                        message=msg,
                                        tab_close=True,
                                        wait_time=self.delay)

    def send(self, phone_no: str, msg: str):
        try:
            whatsapp_thread = threading.Thread(target=self.whatsapp_send, args=(phone_no, msg))
            whatsapp_thread.start()

            self.invalid_account_check(phone_no)
            whatsapp_thread.join()
            print(f"Message sent to {phone_no}!")
        except Exception as err:
            print(f'error: {str(err)}')

    def set_value(self, txt_source_file, float_delay=15):
        self.source_file = txt_source_file
        self.delay = float_delay

    def set_stop(self, status: bool):
        self.stop_sending = status
        print(f'status: {self.stop_sending}')

    def __init__(self, btn_send=None):
        self.source_file = ""
        self.stop_sending = False
        self.delay = 13
        self.btn = btn_send

    def run(self) -> None:
        # in case the application does not exist
        self.set_stop(False)

        df = pd.read_excel(self.source_file, index_col='姓名', dtype={'监护人电话': str})
        time.sleep(3)

        for name, rows in df.iterrows():

            if self.stop_sending:
                print('stopping......')
                messagebox.showinfo(title='已停止发送', message=f'中断于同学：{name}')
                if self.btn is not None:
                    self.btn.configure(text="批量发送信息")
                return

            elif pd.isna(rows['监护人电话']) or rows['离舍'] == 3:
                continue

            elif rows['离舍'] == 0:
                message = f"{name} 同学 {rows['寝室']} 于 {str(rows['日期'])[:10]} 本周留舍 (留校)。\n敬请家长/监护人关注。"
                self.send(rows['监护人电话'], message)

            elif rows['离舍'] == 1:
                text_date = ' (星期' + get_date_of_week(rows['日期']) + ')'
                timing = str(rows['日期'])[:10] + text_date + ' ' + str(rows['时间'])[:5]
                message = f'{name} 同学 {rows['寝室']} 于 {timing} 离校 (回家)。\n敬请家长/监护人关注。'
                self.send(rows['监护人电话'], message)

            elif rows['离舍'] == 2:  # case that when other message is not null
                message = str(rows['其他信息'])
                self.send(rows['监护人电话'], message)

        if self.btn is not None:
            self.btn.configure(text="批量发送信息")


def local():
    a = WebWhatsappSendMessage()
    a.set_value('w.xlsx')
    a.run()


def test():
    friend_numbers = ["+60163490531", "+6586518193", "+60126580830", ]
    message = "Hey bro, have you completed your Python task?"
    send_whatsapp_message(msg=message, phone_numbers=friend_numbers)

"""
if __name__ == "__main__":

    try:
        test()
    except Exception as e:
        print(f'error: {str(e)}')
        exit(0)
        """
