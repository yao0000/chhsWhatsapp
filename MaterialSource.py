import os
import sys


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """

    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def file_path(relative_path):
    return os.getcwd() + '\\_internal\\assets\\' + relative_path


# for GUI
logo = file_path('i.rpr')
required_columns = ['寝室', '姓名', '日期', '时间', '监护人电话', '离舍', '其他信息']

# for Application
textbox_images = ['et1.png',
                  'et2.png',
                  'et3.png',
                  'et4.png']

textbox_images_paths = [file_path(filename) for filename in textbox_images]

# for WebApp
user_not_found_images = ['Invalid_phone_number_eng_light.rpr',
                         'Invalid_phone_number_eng_dark.rpr',
                         'Invalid_phone_number_zh_light.rpr',
                         'Invalid_phone_number_zh_dark.rpr']

user_not_found_images_paths = [file_path(filename) for filename in user_not_found_images]
