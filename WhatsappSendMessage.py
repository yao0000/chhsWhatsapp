import webbrowser

import pyautogui
import time
import os
import sys
import pandas as pd
import pyperclip
from tkinter import messagebox
from pyautogui import size

from MaterialSource import (
    textbox_images_paths,
    user_not_found_images_paths
)

import Log


def get_date_of_week(date):
    switcher = {
        "Monday": "一",
        "Tuesday": "二",
        "Wednesday": "三",
        "Thursday": "四",
        "Friday": "五",
        "Saturday": "六",
        "Sunday": "日"
    }
    day_name = date.strftime("%A")
    return switcher.get(day_name, '')


class WhatsappParentNotice:

    @staticmethod
    def user_not_found(phone_number: str):
        accuracy = 0.8

        # have to modify
        for image in user_not_found_images_paths:
            try:
                coords = pyautogui.locateOnScreen(image, confidence=accuracy)

                # 如果找到图片，点击图片的中心位置
                if coords:
                    Log.log_message(_time=time.localtime(),
                                    receiver=phone_number,
                                    message='用户不存在')
                    return
                else:
                    print('not found')

            except Exception as e:
                print(f'err: {str(e)}')

    @staticmethod
    def acc_click():
        # 设置图片查找的准确度，范围是0-1，1是最准确，但速度慢
        accuracy = 0.8

        for image in textbox_images_paths:
            # 查找图片，返回图片的中心坐标
            try:
                coords = pyautogui.locateOnScreen(image, confidence=accuracy)

                # 如果找到图片，点击图片的中心位置
                if coords:
                    x, y = pyautogui.center(coords)
                    pyautogui.click(x, y)
                    return
                else:
                    print('not found')

            except Exception as e:
                print(f'err: {str(e)}')

    def send_whatsapp_message(self, phone_number, message):
        # whatsapp_url = f"whatsapp://send?phone={phone_number}&text={message}"
        whatsapp_url = f"whatsapp://send?phone={phone_number}"
        webbrowser.open(whatsapp_url)

        # Wait for WhatsApp to open
        time.sleep(self.delay)  # Adjust the wait time as needed
        self.acc_click()
        pyperclip.copy(message)
        pyautogui.click(self.WIDTH * 0.6, self.HEIGHT * 0.9)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')

        # In case the whatsapp number is not existing
        time.sleep(self.delay)

        pyautogui.click(self.WIDTH * 0.6, self.HEIGHT * 0.9)
        pyautogui.click(self.WIDTH * 0.6, self.HEIGHT * 0.9)
        Log.log_message(_time=time.localtime(), receiver=phone_number, message=message)

    def run(self):
        self.set_stop(False)

        df = pd.read_excel(self.source_file, index_col='姓名', dtype={'监护人电话': str})

        for name, rows in df.iterrows():
            if pd.isna(rows['监护人电话']):
                continue
            print(rows['监护人电话'])

            if not pd.isna(rows['其他信息']):
                message = rows['其他信息']
                self.send_whatsapp_message(rows['监护人电话'], message)
                continue

            if rows['离舍'] == 0:
                message = f"{name} 同学 {rows['寝室']} 于 {self.return_date} 本周留舍 (留校)。\n敬请家长/监护人关注。"
                self.send_whatsapp_message(rows['监护人电话'], message)

            elif rows['离舍'] == 1:
                date = ' (星期' + get_date_of_week(rows['日期']) + ')'
                date = str(rows['日期'])[:10] + date + ' ' + str(rows['时间'])[:5]
                message = f'{name} 同学 {rows['寝室']} 于 {date} 离校 (回家)。\n敬请家长/监护人关注。'
                self.send_whatsapp_message(rows['监护人电话'], message)

            if self.stop_sending:
                messagebox.showinfo(message=f'已停止发送。\n最后一位信息接收者：{name}')
                break

    def __init__(self):
        self.source_file = ""
        self.delay = 2
        self.return_date = ""
        self.stop_sending = False

        self.WIDTH, self.HEIGHT = size()

    def set_value(self, txt_source_file, return_date, float_delay=2):
        self.source_file = txt_source_file
        self.delay = float_delay
        self.return_date = return_date

    def set_stop(self, status: bool):
        self.stop_sending = status


if __name__ == '__main__':
    app = WhatsappParentNotice()
    app.set_value("w.xlsx", "2024-04-05")
    print('called in main function')
    time.sleep(1)
    app.run()

