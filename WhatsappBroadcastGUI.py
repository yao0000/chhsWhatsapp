import threading

from PIL import Image
from tkinter import messagebox

import os
import customtkinter
import pandas as pd

from WhatsappSendMessage import WhatsappParentNotice


class WhatsappParentNoticeGUI(customtkinter.CTk):

    def button_trigger_event(self):
        if self.tabview.get() == self.TAB1:
            if self.btn_send_message.cget('text') == self.STRING_BATCH_SEND:
                if not self.display_example_event():
                    return

                self.btn_send_message.configure(text=self.STRING_CANCEL_SEND)

                whatsapp_thread = threading.Thread(target=lambda: (
                    self.whatsapp.set_value(
                        self.string_filepath.get(),
                        float(self.option_delay.get())
                    ),
                    self.whatsapp.run(),
                    self.btn_send_message.configure(text=self.STRING_BATCH_SEND)
                ))
                whatsapp_thread.start()

            elif self.btn_send_message.cget("text") == self.STRING_CANCEL_SEND:
                self.btn_send_message.configure(text=self.STRING_BATCH_SEND)
                self.whatsapp.set_stop(True)

    def display_example_event(self):
        if not self.file_existing_check():
            return False

        df = pd.read_excel(self.string_filepath.get())

        if self.tabview.get() == self.TAB1:
            column_to_check = self.whatsapp.staying_leaving_columns

            if not all(column in df.columns for column in column_to_check):
                messagebox.showerror(title="错误", message=f'Excel文件表头须拥有{column_to_check}\n请查阅"使用须知"')
                return False

            df.set_index('姓名', inplace=True)
            is_displayed_leave = False
            is_display_stay = False

            for name, rows in df.iterrows():

                if rows['离舍'] == 1 and not is_displayed_leave:
                    text_date = ' (星期' + self.whatsapp.get_date_of_week(rows['日期']) + ')'
                    timing = str(rows['日期'])[:10] + text_date + ' ' + str(rows['时间'])[:5]
                    message = f'{name} 同学 {rows['寝室']} 于 {timing} 离校 (回家)。\n敬请家长/监护人关注。'

                    self.string_leaving_sample.set(message)
                    is_displayed_leave = True

                elif rows['离舍'] == 0 and not is_display_stay:
                    message = f"{name} 同学 {rows['寝室']} 于 {str(rows['日期'])[:10]} 本周留舍 (留校)。\n敬请家长/监护人关注。"

                    self.string_staying_sample.set(message)

                if is_display_stay and is_displayed_leave:
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
        self.string_leaving_sample.set("")
        self.string_staying_sample.set("")

    def sidebar_grid_init(self):
        self.sidebar_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")

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

        pic_path = os.getcwd() + "\\_internal\\src\\i.rpr"
        logo = customtkinter.CTkImage(light_image=Image.open(pic_path),
                                      dark_image=Image.open(pic_path),
                                      size=(125, 125))
        label_logo = customtkinter.CTkLabel(self.sidebar_frame, image=logo,
                                            text="")  # display image with a CTkLabel
        label_logo.grid(row=5, column=0, padx=5, pady=5)

    def tab1_grid_init(self):
        self.tabview.tab(self.TAB1).grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.tabview.tab(self.TAB1).grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=2)

        # section 1: file chooser
        label_filepath = customtkinter.CTkLabel(self.tabview.tab(self.TAB1), text="文件路径: ")
        label_filepath.grid(row=0, column=0, padx=10, pady=10)
        entry_filepath = customtkinter.CTkEntry(self.tabview.tab(self.TAB1), textvariable=self.string_filepath,
                                                state="disable")
        entry_filepath.grid(row=0, column=1, columnspan=2, padx=5, pady=10, sticky="ew")
        self.btn_filechooser.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        # section 2: clear or generate sample message
        self.btn_clear_sample.grid(row=1, column=2, padx=10, pady=(20, 10), sticky="ew")
        self.btn_generate_example.grid(row=1, column=3, padx=10, pady=(20, 10), sticky="ew")

        # section 3: To display leaving message
        label_leave_hostel_eg = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text="离舍范本：")
        label_leave_hostel_eg.grid(row=2, column=0, padx=10, pady=10)
        entry_leaving_sample = customtkinter.CTkEntry(master=self.tabview.tab(self.TAB1),
                                                      textvariable=self.string_leaving_sample,
                                                      state="disable")
        entry_leaving_sample.grid(row=2, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

        # section 4: To display staying message
        label_staying_hostel_eg = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text="留舍范本：")
        label_staying_hostel_eg.grid(row=3, column=0, padx=10, pady=10)
        entry_staying_sample = customtkinter.CTkEntry(master=self.tabview.tab(self.TAB1),
                                                      textvariable=self.string_staying_sample,
                                                      state="disable")
        entry_staying_sample.grid(row=3, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

        # section 5: Allow to modify the interval of sending message
        label_delay = customtkinter.CTkLabel(master=self.tabview.tab(self.TAB1), text='信息发送延迟间隔：')
        label_delay.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
        self.option_delay.grid(row=4, column=2, padx=10, pady=10, sticky="ew")

        # section 6: To send message in batch
        self.btn_send_message.grid(row=5, column=3, padx=10, pady=10, sticky="ew")

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
                ('A1234', '学生姓名', '2020-01-01', '17:30:00', '0123456789', '1')]

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
                                                 text="注意事项: \n"
                                                      "1. 日期格式为: 2020-01-01 (YYYY-MM-DD)\n"
                                                      "2. 时间格式为24小时制 HH:mm:ss (例 18:30:00)\n"
                                                      "3. 监护人电话-马来西亚号码以外都需要加上区号"
                                                      " (例，新加坡 +65)\n"
                                                      "4. 离舍使用数字 '1'(离舍) \\ '0'(留舍) 区别",
                                                 anchor="nw",
                                                 justify="left")
        label_attention.grid(row=3, column=0, padx=5, columnspan=3, pady=(20, 10), sticky="nw")

    def config(self):
        # configure the attribute
        # tab 1 attributes
        self.btn_filechooser.configure(command=self.filechooser_event)
        self.btn_generate_example.configure(command=self.display_example_event)
        self.btn_clear_sample.configure(command=self.clear_string_sample)
        self.btn_send_message.configure(command=self.button_trigger_event)
        self.option_delay.set('1')
        self.tabview.set(self.TAB1)
        self.string_filepath.set("./w.xlsx")
        self.scaling_option_menu.set("100%")

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
        self.geometry(f"{900}x{380}")
        self.change_appearance_mode_event("深色模式")  # set default appearance

        # backend whatsapp message sending initialisation
        self.whatsapp = WhatsappParentNotice()

        # configure grid layout
        self.grid_columnconfigure(1, weight=2)
        self.grid_columnconfigure((0, 1, 2, 3), weight=2)
        self.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)

        # sidebar init
        self.sidebar_frame = customtkinter.CTkFrame(self, width=250, corner_radius=20)
        self.appearance_mode_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                                       values=["深色模式", "亮色模式", "系统"],
                                                                       command=self.change_appearance_mode_event)
        self.scaling_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                               values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.sidebar_grid_init()

        # tabview init
        self.tabview = customtkinter.CTkTabview(self, anchor="nw")
        self.tabview.grid(row=0, column=1, rowspan=7, columnspan=4, padx=10, pady=10, sticky="ew")

        self.TAB1 = "监护人通知"
        self.TAB3 = "使用须知"
        self.tabview.add(self.TAB1)
        self.tabview.add(self.TAB3)

        # tab 1 attributes initialization
        self.string_filepath = customtkinter.StringVar()
        self.btn_filechooser = customtkinter.CTkButton(master=self.tabview.tab(self.TAB1), text="打开文件",
                                                       fg_color="transparent", border_width=2)
        self.btn_clear_sample = customtkinter.CTkButton(self.tabview.tab(self.TAB1), text="清除")
        self.btn_generate_example = customtkinter.CTkButton(self.tabview.tab(self.TAB1), text="生成范本")
        self.string_leaving_sample = customtkinter.StringVar()
        self.string_staying_sample = customtkinter.StringVar()
        self.option_delay = customtkinter.CTkOptionMenu(master=self.tabview.tab(self.TAB1),
                                                        dynamic_resizing=False,
                                                        values=['0.5', '1', '1.5', '2', '2.5',
                                                                '3', '3.5', '4', '4.5', '5'])

        self.STRING_BATCH_SEND = "批量发送信息"
        self.STRING_CANCEL_SEND = "中断发送"
        self.btn_send_message = customtkinter.CTkButton(master=self.tabview.tab(self.TAB1),
                                                        text=self.STRING_BATCH_SEND)
        # end of tab 1

        # tab 3 attribute
        self.scrollable_frame = customtkinter.CTkScrollableFrame(master=self.tabview.tab(self.TAB3), width=600)

        self.tab1_grid_init()
        # self.tab3_grid_init()

        self.config()


if __name__ == "__main__":
    try:
        app = WhatsappParentNoticeGUI()
        app.mainloop()
    except Exception as error:
        messagebox.showerror(message=str(error))
