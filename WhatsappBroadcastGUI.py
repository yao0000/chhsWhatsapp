import datetime
import os
import threading
import tkinter
import webbrowser
from tkinter import messagebox, font

import customtkinter
import pandas as pd
from CTkMessagebox import CTkMessagebox
from PIL import Image
from tkcalendar import Calendar as DatePicker

from WebWhatsappSendMessage import WebWhatsappSendMessage
from WhatsappSendMessage import WhatsappParentNotice
from WhatsappSendMessage import get_date_of_week

from MaterialSource import (
    logo,
    required_columns
)

import time
start_time = 0


class WhatsappParentNoticeGUI(customtkinter.CTk):

    def update_option(self):

        if self.mode_var.get() == self.WEB_APP:
            self.option_delay.configure(values=self.OPTION_WEB)
            self.option_delay.set(self.OPTION_WEB[1])
        elif self.mode_var.get() == self.INSTALLED_APP:
            self.option_delay.configure(values=self.OPTION_APP)
            self.option_delay.set(self.OPTION_APP[1])

    def button_trigger_event(self):
        if self.tabview.get() == self.TAB1:

            if self.btn_send_message.cget('text') == self.STRING_BATCH_SEND:
                if self.mode_var.get() == self.INSTALLED_APP:
                    """
                                        respond = CTkMessagebox(title="提示",
                                                                message='开始作业前, 确保已登入。\n请点击\nwhatsapp聊天室的文字输入框。',
                                                                icon='info',
                                                                option_1='取消',
                                                                option_2='开始作业')

                                        if respond.get() == '取消' or respond.get() is None:
                                            return

                                        webbrowser.open('whatsapp://')
                                        CTkMessagebox(title='提示',
                                                      message='程序执行中，勿操作电脑。\n中断作业除外。',
                                                      icon='info',
                                                      option_1='开始').get()

                                        self.btn_send_message.configure(text=self.STRING_CANCEL_SEND)

                                        self.whatsapp_thread = threading.Thread(target=lambda: (
                                            self.whatsapp.set_value(
                                                self.string_filepath.get(),
                                                self.string_return_date.get(),
                                                float(self.option_delay.get())
                                            ),
                                            self.whatsapp.run(),
                                        ))
                                        self.whatsapp_thread.start()
                                        """
                    CTkMessagebox(title='提示',
                                  message='尚未开发完成。请联系负责人。',
                                  icon='info',
                                  option_1='确认').get()
                    return

                elif self.mode_var.get() == self.WEB_APP:
                    webbrowser.open('https://web.whatsapp.com/')
                    respond = CTkMessagebox(title="提示",
                                            message='请确保默认浏览器已登入\nWhatsApp账号。\n并最大化浏览器窗口。',
                                            icon='info',
                                            option_1='取消',
                                            option_2='开始作业')

                    if respond.get() == '取消' or respond.get() is None:
                        return

                    CTkMessagebox(title='提示',
                                  message='程序执行中，勿操作电脑。\n中断作业除外。',
                                  icon='info',
                                  option_1='确认').get()

                    self.btn_send_message.configure(text=self.STRING_CANCEL_SEND)
                    self.webapp_thread = threading.Thread(target=lambda: (
                        self.web_app.set_value(
                            self.string_filepath.get(),
                            int(self.option_delay.get())
                        ),
                        self.web_app.run()
                    ))
                    self.webapp_thread.start()

            elif self.btn_send_message.cget("text") == self.STRING_CANCEL_SEND:
                self.btn_send_message.configure(text='中断中...')
                self.whatsapp.set_stop(True)
                self.web_app.set_stop(True)

    def display_example_event(self):
        if not self.file_existing_check():
            return False

        df = pd.read_excel(self.string_filepath.get())

        if self.tabview.get() == self.TAB1:
            if not all(column in df.columns for column in required_columns):
                messagebox.showerror(title="错误",
                                     message=f'Excel文件列名 (column) 须含有{required_columns}\n请查阅"使用说明"')
                return False

            df.set_index('姓名', inplace=True)
            is_displayed_leave = False
            is_displayed_staying = False

            for name, rows in df.iterrows():

                if rows['离舍'] == 1 and not is_displayed_leave:
                    text_date = ' (星期' + get_date_of_week(rows['日期']) + ')'
                    timing = str(rows['日期'])[:10] + text_date + ' ' + str(rows['时间'])[:5]
                    message = f'{name} 同学 {rows['寝室']} 于 {timing} 离校 (回家)。\n敬请家长/监护人关注。'

                    self.textbox_leaving.delete(1.0, tkinter.END)
                    self.textbox_leaving.insert('0.0', message)
                    is_displayed_leave = True

                elif rows['离舍'] == 0 and not is_displayed_staying:
                    message = f"{name} 同学 {rows['寝室']} 于 {str(rows['日期'])[:10]} 本周留舍 (留校)。\n敬请家长/监护人关注。"
                    self.textbox_staying.delete(1.0, tkinter.END)
                    self.textbox_staying.insert('0.0', message)
                    is_displayed_staying = True

                if is_displayed_leave and is_displayed_staying:
                    break

        return True

    def filechooser_event(self):
        file_path = customtkinter.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.tabview.get() == self.TAB1:
            self.string_filepath.set(file_path)

    # To check whether the file existing
    def file_existing_check(self):
        src_file = ""
        if self.tabview.get() == self.TAB1:
            src_file = self.string_filepath.get()

        if src_file == "":
            messagebox.showerror(title="错误", message="请选择源文件")
            return False

        if not os.path.isfile(src_file):
            messagebox.showerror(title="错误", message="文件不存在")
            return False

        return True

    @staticmethod
    def change_scaling_event(new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def clear_string_sample(self):
        self.textbox_leaving.delete(1.0, tkinter.END)
        self.textbox_staying.delete(1.0, tkinter.END)

    def sidebar_grid_init(self):
        self.sidebar_frame.grid(row=0, column=0, rowspan=6, sticky="e")

        label_program = customtkinter.CTkLabel(self.sidebar_frame, text="离舍/留宿通知系统",
                                               font=customtkinter.CTkFont(size=20, weight="bold"),
                                               anchor="center", justify="center")
        label_program.grid(row=0, column=0, padx=(10, 10), pady=(10, 10), sticky="ew")

        appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="外观模式:", anchor="w")
        appearance_mode_label.grid(row=1, column=0, padx=5, pady=5)

        scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="分辨率:", anchor="w")
        scaling_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_option_menu.grid(row=2, column=0, padx=5, pady=5)
        self.scaling_option_menu.grid(row=4, column=0, padx=5, pady=5)

        img_path = logo
        img_logo = customtkinter.CTkImage(light_image=Image.open(img_path),
                                          dark_image=Image.open(img_path),
                                          size=(125, 125))
        label_logo = customtkinter.CTkLabel(self.sidebar_frame, image=img_logo,
                                            text="")  # display image with a CTkLabel
        label_logo.grid(row=5, column=0, padx=5, pady=5)

    def tab1_grid_init(self):
        self.tabview.tab(self.TAB1).grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.tabview.tab(self.TAB1).grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=2)

        # section 1: file chooser
        label_filepath = customtkinter.CTkLabel(self.tabview.tab(self.TAB1), text="文件路径：")
        label_filepath.grid(row=0, column=0, padx=10, pady=10)
        entry_filepath = customtkinter.CTkEntry(self.tabview.tab(self.TAB1), textvariable=self.string_filepath,
                                                state="disable")
        entry_filepath.grid(row=0, column=1, columnspan=2, padx=5, pady=10, sticky="ew")
        self.btn_filechooser.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        # date selection
        label_date = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text='默认离舍日：')
        label_date.grid(row=1, column=0, padx=10, pady=10)
        entry_date = customtkinter.CTkEntry(master=self.tabview.tab(self.TAB1), textvariable=self.string_return_date,
                                            state='disable')
        entry_date.grid(row=1, column=1, columnspan=2, padx=5, pady=10, sticky='ew')
        self.btn_date.grid(row=1, column=3, padx=10, pady=10, sticky='ew')

        # section 2: clear or generate sample message
        self.btn_clear_sample.grid(row=2, column=2, padx=10, pady=(20, 5), sticky="ew")
        self.btn_generate_example.grid(row=2, column=3, padx=10, pady=(20, 5), sticky="ew")

        # section 3: To display leaving message
        label_leave_hostel_eg = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text="离舍范本：")
        label_leave_hostel_eg.grid(row=3, column=0, padx=10, pady=5)
        self.textbox_leaving.grid(row=3, column=1, columnspan=3, padx=10, pady=5, sticky='ew')

        # belongs to display section
        label_stay_hostel_eg = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text='留宿范本：')
        label_stay_hostel_eg.grid(row=4, column=0, padx=10, pady=(5, 10))
        self.textbox_staying.grid(row=4, column=1, columnspan=3, padx=10, pady=(5, 10), sticky='ew')

        # section 4: mode selection
        label_mode = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text="运行平台：")
        label_mode.grid(row=5, column=3, padx=10, pady=(10, 2), sticky='nw')
        self.radio_web.grid(row=6, column=3, padx=1, pady=1, sticky="ew")
        self.radio_app.grid(row=7, column=3, padx=1, pady=1, sticky="ew")

        # section 5: Allow to modify the interval of sending message
        label_delay = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1),
                                             text='信息发送延迟间隔：',
                                             anchor='nw')
        label_delay.grid(row=6, column=0, columnspan=2, padx=20, pady=1, sticky='ew')
        self.option_delay.grid(row=7, column=0, padx=(20, 0), pady=1, sticky='ns')
        label_second = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text='(秒)')
        label_second.grid(row=7, column=1, padx=1, pady=1, sticky="nw")

        # section 6: To send message in batch
        self.btn_send_message.grid(row=8, column=3, padx=10, pady=(10, 0), sticky="ew")

    def tab3_grid_init(self):
        self.tabview.tab(self.TAB3).grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.tabview.tab(self.TAB3).grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=2)
        self.scrollable_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # section 1: labelling
        label_title1 = customtkinter.CTkLabel(master=self.scrollable_frame, text='Excel文件格式样本: ',
                                              anchor="nw")
        label_title1.grid(row=0, column=0, columnspan=3, padx=1, pady=1, sticky="ew")

        # section 2: data sample display
        data = [('寝室', '姓名', '日期', '时间', '监护人电话', '离舍'),
                ('A1234', '学生姓名', '2020-01-01', '17:30:00', '+60123456789', '1')]

        data_rows_count = len(data)
        data_columns_count = len(data[0])

        for i in range(data_rows_count):
            for j in range(data_columns_count):
                d = customtkinter.CTkLabel(master=self.scrollable_frame,
                                           corner_radius=10,
                                           text_color='black',
                                           fg_color='grey',
                                           text=data[i][j])
                d.grid(row=(i + 1), column=j, sticky="ew")

        # section 3 format information
        label_attention = customtkinter.CTkLabel(master=self.scrollable_frame,
                                                 text="宿舍通知注意事项: \n"
                                                      "- 日期格式为: 2020-01-01 (YYYY-MM-DD)\n"
                                                      "- 时间格式为24小时制 HH:mm:ss (例 18:30:00)\n"
                                                      "- 监护人电话：需加上区号 (+6)\n"
                                                      "- 离舍使用数字 '1'(离舍) \\ '0'(留舍) 区别",
                                                 anchor="nw",
                                                 justify="left")
        label_attention.grid(row=3, column=0, padx=5, columnspan=3, pady=(20, 10), sticky="nw")

    def open_date_picker(self):
        self.date_picker_window = tkinter.Toplevel(self)
        self.date_picker_window.title('选择日期')
        custom_font = font.Font(family='Times New Roman', size=18)
        self.date_picker = DatePicker(self.date_picker_window,
                                      selectmode='day',
                                      date_pattern='yyyy-mm-dd',
                                      font=custom_font)
        self.btn_date_picker = tkinter.Button(self.date_picker_window,
                                              text='确认',
                                              font=custom_font,
                                              command=self.confirm_date)
        self.date_picker.pack(padx=10, pady=10)
        self.btn_date_picker.pack(padx=10, pady=10)

    def confirm_date(self):
        selected_date = self.date_picker.selection_get()
        selected_date = str(selected_date) + ' (星期' + get_date_of_week(selected_date) + ')'
        self.string_return_date.set(selected_date)
        self.date_picker_window.destroy()

    def config(self):
        # configure the attribute
        self.btn_filechooser.configure(command=self.filechooser_event)
        self.btn_date.configure(command=self.open_date_picker)
        self.btn_generate_example.configure(command=self.display_example_event)
        self.btn_clear_sample.configure(command=self.clear_string_sample)
        self.btn_send_message.configure(command=self.button_trigger_event)
        self.radio_web.configure(command=self.update_option)
        self.radio_app.configure(command=self.update_option)
        self.mode_var.set(self.WEB_APP)
        self.option_delay.set(self.OPTION_WEB[2])
        self.option_delay.configure(values=self.OPTION_WEB)
        self.tabview.set(self.TAB1)
        self.string_filepath.set("./w.xlsx")
        self.scaling_option_menu.set("100%")

        # set today date
        today = datetime.date.today()
        formatted_date = today.strftime('%Y-%m-%d')
        today_datetime = datetime.datetime.strptime(formatted_date, '%Y-%m-%d')
        formatted_today = f"{formatted_date} (星期{get_date_of_week(today_datetime)})"
        self.string_return_date.set(formatted_today)

    @staticmethod
    def change_appearance_mode_event(new_appearance_mode: str):
        """
        Change system appearance mode
        :param new_appearance_mode: ["Dark", "Light", "System"]
        """
        if new_appearance_mode == "深色模式":
            customtkinter.set_appearance_mode("Dark")
        elif new_appearance_mode == "亮色模式":
            customtkinter.set_appearance_mode("Light")
        else:
            customtkinter.set_appearance_mode("System")

    def __init__(self):
        super().__init__()

        self.title("Whatsapp 通知系统")
        self.geometry(f"{850}x{450}")
        self.change_appearance_mode_event("深色模式")  # set default appearance
        self.minsize(850, 450)

        # configure grid layout
        self.grid_columnconfigure(1, weight=2)
        self.grid_columnconfigure((0, 1, 2, 3), weight=2)
        self.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)

        # sidebar init
        self.sidebar_frame = customtkinter.CTkFrame(self, width=250, corner_radius=20)
        self.appearance_mode_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                                       values=["深色模式", "亮色模式", "系统"],
                                                                       command=self.change_appearance_mode_event,
                                                                       anchor='nw')
        self.scaling_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                               values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.sidebar_grid_init()

        # tabview init
        self.tabview = customtkinter.CTkTabview(self, anchor="nw")
        self.tabview.grid(row=0, column=1, rowspan=7, columnspan=4, padx=10, pady=10)

        self.TAB1 = "监护人通知"
        self.TAB3 = "使用须知"
        self.tabview.add(self.TAB1)
        self.tabview.add(self.TAB3)

        # date picker init
        self.date_picker_window = None
        self.date_picker = None
        self.btn_date_picker = None

        # tab 1 attributes initialization
        self.string_filepath = customtkinter.StringVar()
        self.btn_filechooser = customtkinter.CTkButton(master=self.tabview.tab(self.TAB1),
                                                       text="打开文件",
                                                       border_width=2)

        self.string_return_date = customtkinter.StringVar()
        self.btn_date = customtkinter.CTkButton(master=self.tabview.tab(self.TAB1),
                                                text='选择日期',
                                                border_width=2)

        self.btn_clear_sample = customtkinter.CTkButton(self.tabview.tab(self.TAB1),
                                                        text="清除信息",
                                                        border_width=2)
        self.btn_generate_example = customtkinter.CTkButton(self.tabview.tab(self.TAB1),
                                                            text="文件读取测试",
                                                            border_width=2)
        self.textbox_leaving = customtkinter.CTkTextbox(master=self.tabview.tab(self.TAB1),
                                                        height=50)
        self.textbox_staying = customtkinter.CTkTextbox(master=self.tabview.tab(self.TAB1),
                                                        height=50)

        self.OPTION_APP = ['1', '2', '3', '4', '5']
        self.OPTION_WEB = ['8', '10', '12', '14', '16']
        self.option_delay = customtkinter.CTkOptionMenu(master=self.tabview.tab(self.TAB1),
                                                        dynamic_resizing=False)

        self.STRING_BATCH_SEND = "批量发送信息"
        self.STRING_CANCEL_SEND = "中断发送"
        self.btn_send_message = customtkinter.CTkButton(master=self.tabview.tab(self.TAB1),
                                                        text=self.STRING_BATCH_SEND,
                                                        border_width=2)

        # backend whatsapp message sending initialisation
        self.whatsapp = WhatsappParentNotice()
        self.web_app = WebWhatsappSendMessage(self.btn_send_message)
        self.WEB_APP = '网页版(推荐)'
        self.INSTALLED_APP = 'WhatsApp应用'
        self.mode_var = customtkinter.StringVar()
        self.whatsapp_thread = threading.Thread(target=lambda: (
            self.whatsapp.set_value(
                self.string_filepath.get(),
                self.string_return_date.get(),
                float(self.option_delay.get())
            ),
            self.whatsapp.run(),
        ))

        self.webapp_thread = threading.Thread(target=lambda: (
            self.web_app.set_value(
                self.string_filepath.get(),
                int(self.option_delay.get())
            ),
            self.web_app.run()
        ))

        self.radio_web = customtkinter.CTkRadioButton(master=self.tabview.tab(self.TAB1),
                                                      text=self.WEB_APP,
                                                      variable=self.mode_var,
                                                      value=self.WEB_APP)
        self.radio_app = customtkinter.CTkRadioButton(master=self.tabview.tab(self.TAB1),
                                                      text=self.INSTALLED_APP,
                                                      variable=self.mode_var,
                                                      value=self.INSTALLED_APP)
        # end of tab 1

        # tab 3 attribute
        self.scrollable_frame = customtkinter.CTkScrollableFrame(master=self.tabview.tab(self.TAB3), width=600)

        self.tab1_grid_init()
        self.tab3_grid_init()

        self.config()
        end_time = time.time()
        print(f'duration: {end_time - start_time} seconds')


if __name__ == "__main__":
    start_time = time.time()
    app = WhatsappParentNoticeGUI()
    try:
        app.mainloop()
    except Exception as error:
        messagebox.showerror(message=str(error))
        app.destroy()
