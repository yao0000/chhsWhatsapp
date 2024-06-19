import os
import time
from datetime import datetime
from MaterialSource import resource_path

date = datetime.today().date()


def format_message(message: str) -> str:
    """Formats the Message to remove redundant Spaces and Newline chars"""
    msg_l = message.split(" ")
    new = []
    for x in msg_l:
        if "\n" in x:
            x = x.replace("\n", "")
            new.append(x) if not len(x) == 0 else None

        elif len(x) != 0:
            new.append(x)

    return " ".join(new)


def log_message(_time: time.struct_time, receiver: str, message: str, floor: str, name: str) -> None:
    """Logs the Message Information after it is Sent"""

    filename = f'错误日志_{date}.txt'
    directory = os.getcwd() + "\\" + filename
    if not os.path.exists(directory):
        file = open(directory, "w+")
        file.close()

    message = format_message(message)

    with open(directory, "a", encoding="utf-8") as file:
        file.write(
            f"日期: {_time.tm_mday}/{_time.tm_mon}/{_time.tm_year}\n时间: {_time.tm_hour}:{_time.tm_min}\n"
            f"姓名: {name}\n寝室: {floor}\n"
            f"监护人电话: {receiver}\n信息: {message}"
        )
        file.write("\n--------------------\n")
        file.close()
