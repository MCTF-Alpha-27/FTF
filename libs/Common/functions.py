"""
函数库
"""

import os
import time
import logging

from colorama import Fore, init
from libs.Common import config

logging.basicConfig(
    filename=f"logs/{time.strftime(r'%Y-%m-%d-%H.%M.%S')}.log",
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s]: %(message)s",
    encoding="utf-8",
)
init()

if not os.path.exists("logs"):
    os.mkdir("logs")


def log(text, level="normal", *, logfile_only=False):
    if level == "normal":
        print(Fore.LIGHTGREEN_EX + text)
    elif level == "info":
        if not logfile_only:
            print(
                Fore.LIGHTGREEN_EX
                + f"[{time.strftime(r'%Y-%m-%d %H:%M:%S')}] [INFO]: {text}"
            )
        logging.info(text)
    elif level == "warning":
        if not logfile_only:
            print(
                Fore.LIGHTYELLOW_EX
                + f"[{time.strftime(r'%Y-%m-%d %H:%M:%S')}] [WARNING]: {text}"
            )
        logging.warning(text)
        print(Fore.LIGHTGREEN_EX, end="")
    elif level == "error":
        if not logfile_only:
            print(
                Fore.LIGHTRED_EX
                + f"[{time.strftime(r'%Y-%m-%d %H:%M:%S')}] [ERROR]: {text}"
            )
        logging.error(text)
        print(Fore.LIGHTGREEN_EX, end="")
    elif level == "exception":
        if not logfile_only:
            print(
                Fore.LIGHTRED_EX
                + f"[{time.strftime(r'%Y-%m-%d %H:%M:%S')}] [ERROR]: {text}"
            )
        logging.exception(text)
        print(Fore.LIGHTGREEN_EX, end="")
    elif level == "debug":
        if config.debug:
            if not logfile_only:
                print(
                    Fore.LIGHTBLUE_EX
                    + f"[{time.strftime(r'%Y-%m-%d %H:%M:%S')}] [DEBUG]: {text}"
                )
        logging.debug(text)
        print(Fore.LIGHTGREEN_EX, end="")
    else:
        if not logfile_only:
            print(f"[{time.strftime(r'%Y-%m-%d %H:%M:%S')}] [{level}]: {text}")
        logging.info(text)
        print(Fore.LIGHTGREEN_EX, end="")


def choice(choose="YN", text="Y/N", default=None, timeout=10, *, hide=False):
    choice = "choice /C "
    if " " in text:
        raise SyntaxError("显示的文字中不能含有空格")
    if hide:
        choice = choice + choose + " /N " + " /M " + text
    else:
        choice = choice + choose + " /M " + text
    if default:
        if default in choose:
            choice = choice + " /D " + default
            choice = choice + " /T " + str(timeout)
        else:
            raise SyntaxError("按键默认值不在设置的按键中")
    return os.system(choice)


def display_width(s):
    return sum(2 if ord(c) > 127 else 1 for c in s)


def pad(s, width):
    cur = display_width(s)
    return s + " " * (width - cur)


log("初始化终端", "info")
