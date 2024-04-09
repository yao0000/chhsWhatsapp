import webbrowser
import pyautogui
import time
import pandas as pd
import pyperclip


class WhatsappParentNotice:

    def send_whatsapp_message(self, phone_number, message):
        # whatsapp_url = f"whatsapp://send?phone={phone_number}&text={message}"
        whatsapp_url = f"whatsapp://send?phone={phone_number}&text=\"\""
        webbrowser.open(whatsapp_url)

        # Wait for WhatsApp to open
        time.sleep(self.delay)  # Adjust the wait time as needed
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(self.delay)
        pyautogui.press('enter')

        # Wait for the message to be sent
        time.sleep(self.delay)  # Adjust the wait time as needed

    def run(self):
        df = pd.read_excel(self.source_file, index_col='姓名')

        for name, rows in df.iterrows():
            if rows['离舍'] == 0:
                message = f"{name} 同学 {rows['寝室']} 于 {str(rows['日期'])[:10]} 本周留舍 (留校)。\n敬请家长/监护人关注。"
                self.send_whatsapp_message(rows['监护人电话'], message)

            elif rows['离舍'] == 1:
                text_date = ' (星期' + self.get_date_of_week(rows['日期']) + ')'
                timing = str(rows['日期'])[:10] + text_date + ' ' + str(rows['时间'])[:5]
                message = f'{name} 同学 {rows['寝室']} 于 {timing} 离校 (回家)。\n敬请家长/监护人关注。'
                self.send_whatsapp_message(rows['监护人电话'], message)


    @staticmethod
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

    def __init__(self, txt_source_file, float_delay=1):
        self.source_file = txt_source_file
        self.delay = float_delay


if __name__ == '__main__':
    app = WhatsappParentNotice('w.xlsx')
    app.run()
