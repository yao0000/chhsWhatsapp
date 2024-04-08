from PIL import Image
from tkinter import messagebox

import os
import customtkinter
import pandas as pd

import WhatsappSendMessage


class WhatsappParentNoticeGUI(customtkinter.CTk):

    def display_example_event(self):
        if not self.file_validate():
            return

        df = pd.read_excel(self.string_filepath.get())
        column_to_check = ['姓名', '日期', '时间', '手机号', '离舍']

        if not all(column in df.columns for column in column_to_check):
            messagebox.showerror(title="错误", message=f'Excel文件表头须拥有{column_to_check}')
            return

        df.set_index('姓名', inplace=True)
        is_displayed_leave = False
        is_display_stay = False

        for name, rows in df.iterrows():
            if is_display_stay and is_displayed_leave:
                break

            if rows['离舍'] == 1 and not is_displayed_leave:
                text_date = ' (星期' + WhatsappSendMessage.WhatsappParentNotice.get_date_of_week(rows['日期']) + ')'
                timing = str(rows['日期'])[:10] + text_date + ' ' + str(rows['时间'])[:5]
                message = f'{name} 同学将于 {timing} 离校 (回家)。'

                self.string_leaving_sample.set(message)
                is_displayed_leave = True

            elif rows['离舍'] == 0 and not is_display_stay:
                message = f"{name} 同学于 {str(rows['日期'])[:10]} 本周留舍 (留校)。"

                self.string_staying_sample.set(message)

    def filechooser_event(self):
        file_path = customtkinter.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.tabview.get() == self.tab1:
            self.string_filepath.set(file_path)

    def file_validate(self):
        src_file = self.string_filepath.get().strip()
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
        self.geometry(f"{800}x{380}")
        self.change_appearance_mode_event("深色模式")  # set default appearance

        # configure grid layout
        self.grid_columnconfigure(1, weight=2)
        self.grid_columnconfigure((0, 1, 2, 3), weight=2)
        self.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=20)
        self.sidebar_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")
        self.label_program = customtkinter.CTkLabel(self.sidebar_frame, text="离舍/留宿通知系统",
                                                    font=customtkinter.CTkFont(size=20, weight="bold"),
                                                    anchor="center", justify="center")
        self.label_program.grid(row=0, column=0, padx=(10, 10), pady=(10, 10), sticky="ew")

        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="外观模式:", anchor="w")
        self.appearance_mode_label.grid(row=1, column=0, padx=5, pady=5)
        self.appearance_mode_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                                       values=["深色模式", "亮色模式", "系统"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_option_menu.grid(row=2, column=0, padx=5, pady=5)
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="分辨率:", anchor="w")
        self.scaling_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.scaling_option_menu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                               values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_option_menu.grid(row=4, column=0, padx=5, pady=5)
        self.scaling_option_menu.set("100%")

        pic_path = os.getcwd() + "\\_internal\\src\\i.rpr"
        self.logo = customtkinter.CTkImage(light_image=Image.open(pic_path),
                                           dark_image=Image.open(pic_path),
                                           size=(125, 125))
        self.label_logo = customtkinter.CTkLabel(self, image=self.logo, text="")  # display image with a CTkLabel
        self.label_logo.grid(row=5, column=0, padx=5, pady=5)

        self.tabview = customtkinter.CTkTabview(self, anchor="nw")
        self.tab1 = "回家/留宿通知"
        self.tab2 = "关于"
        self.tabview.add(self.tab1)
        self.tabview.grid(row=0, column=1, rowspan=7, columnspan=4, padx=10, pady=10, sticky="ew")
        self.tabview.tab(self.tab1).grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.tabview.tab(self.tab1).grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=2)

        # section 1: file selection
        self.label_filepath = customtkinter.CTkLabel(self.tabview.tab(self.tab1), text="文件路径: ")
        self.label_filepath.grid(row=0, column=0, padx=10, pady=10)

        self.string_filepath = customtkinter.StringVar()
        self.string_filepath.set("./w.xlsx")
        self.entry_filepath = customtkinter.CTkEntry(self.tabview.tab(self.tab1), textvariable=self.string_filepath,
                                                     state="disable")
        self.entry_filepath.grid(row=0, column=1, columnspan=2, padx=5, pady=10, sticky="ew")

        self.btn_filechooser = customtkinter.CTkButton(master=self.tabview.tab(self.tab1), text="打开文件",
                                                       fg_color="transparent", border_width=2)
        self.btn_filechooser.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        # section 2: add a button to generate sample
        self.btn_generate_example = customtkinter.CTkButton(self.tabview.tab(self.tab1), text="生成范本")
        self.btn_generate_example.grid(row=1, column=3, padx=10, pady=(20, 10), sticky="ew")

        # section 3: display leaving sample
        self.label_leave_hostel_eg = customtkinter.CTkLabel(master=self.tabview.tab(self.tab1), text="离舍范本：")
        self.label_leave_hostel_eg.grid(row=2, column=0, padx=10, pady=10)
        self.string_leaving_sample = customtkinter.StringVar()
        self.entry_leaving_sample = customtkinter.CTkEntry(master=self.tabview.tab(self.tab1),
                                                   textvariable=self.string_leaving_sample,
                                                   state="disable")
        self.entry_leaving_sample.grid(row=2, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

        # section 4: display staying message
        self.label_staying_hostel_eg = customtkinter.CTkLabel(master=self.tabview.tab(self.tab1), text="留舍范本：")
        self.label_staying_hostel_eg.grid(row=3, column=0, padx=10, pady=10)
        self.string_staying_sample = customtkinter.StringVar()
        self.entry_staying_sample = customtkinter.CTkEntry(master=self.tabview.tab(self.tab1),
                                                          textvariable=self.string_staying_sample,
                                                          state="disable")
        self.entry_staying_sample.grid(row=3, column=1, columnspan=3, padx=10, pady=10, sticky="ew")

        # section 5: attribute adjustment
        self.label_delay = customtkinter.CTkLabel(master=self.tabview.tab(self.tab1), text='信息发送延迟间隔：')
        self.label_delay.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
        self.option_delay = customtkinter.CTkOptionMenu(master=self.tabview.tab(self.tab1),
                                                        dynamic_resizing=False,
                                                        values=['0.5', '1', '1.5', '2', '2.5',
                                                                '3', '3.5', '4', '4.5', '5'])
        self.option_delay.grid(row=4, column=2, padx=10, pady=10, sticky="ew")

        # section 6: button to send the whatsapp message
        self.btn_send_message = customtkinter.CTkButton(master=self.tabview.tab(self.tab1), text='批量发送信息')
        self.btn_send_message.grid(row=5, column=3, padx=10, pady=10, sticky="ew")


        # configure the attribute
        self.btn_filechooser.configure(command=self.filechooser_event)
        self.btn_generate_example.configure(command=self.display_example_event)
        self.option_delay.set('1')


if __name__ == "__main__":
    app = WhatsappParentNoticeGUI()
    app.mainloop()
