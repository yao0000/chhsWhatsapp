from PIL import Image
from tkinter import messagebox

import os
import customtkinter
import pandas as pd


class WhatsappBroadcast(customtkinter.CTk):

    def display_example_event(self):
        if not self.file_validate():
            return

        df = pd.read_excel(self.string_filepath.get())
        column_to_check = ['姓名', '日期', '时间', '手机号', '离舍']

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
        self.label_program = customtkinter.CTkLabel(self.sidebar_frame, text="学生离校家长通知系统",
                                                    font=customtkinter.CTkFont(size=20, weight="bold"))
        self.label_program.grid(row=0, column=0, padx=(10, 10), pady=(10, 10))

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

        self.tabview = customtkinter.CTkTabview(self)
        self.tab1 = "回家/留宿通知"
        self.tabview.add(self.tab1)
        self.tabview.grid(row=0, column=1, columnspan=4, padx=10, pady=10, sticky="ew")
        self.tabview.tab(self.tab1).grid_columnconfigure((0, 1, 2, 3), weight=1)

        # section 1: file selection
        self.label_filepath = customtkinter.CTkLabel(self.tabview.tab(self.tab1), text="文件路径: ")
        self.label_filepath.grid(row=0, column=0, padx=10, pady=10)

        self.string_filepath = customtkinter.StringVar()
        self.string_filepath.set("./回家通知列表.xlsx")
        self.entry_filepath = customtkinter.CTkEntry(self.tabview.tab(self.tab1), textvariable=self.string_filepath,
                                                     state="disable")
        self.entry_filepath.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        self.btn_filechooser = customtkinter.CTkButton(master=self.tabview.tab(self.tab1), text="打开文件",
                                                       fg_color="transparent", border_width=2)
        self.btn_filechooser.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        # section 2
        self.label_verify = customtkinter.CTkLabel(self.tabview.tab(self.tab1), text="范本：")
        self.label_verify.grid(row=1, column=0, padx=10, pady=10)


        # configure the attribute
        self.btn_filechooser.configure(command=self.filechooser_event)


if __name__ == "__main__":
    app = WhatsappBroadcast()
    app.mainloop()
