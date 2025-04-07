from psutil import process_iter
from colorama import Fore, init
from . import config
import os
import time
import logging
import pyttsx3
import pyautogui
import subprocess

logging.basicConfig(filename="logs/%s.log"%time.strftime(r"%Y-%m-%d-%H.%M.%S"), level=logging.DEBUG, format="[%(asctime)s] [%(levelname)s]: %(message)s", encoding="utf-8")
init()

if not os.path.exists("logs"):
    os.mkdir("logs")

engine = pyttsx3.init()

def say_in_english(words):
    if not config.voice:
        return
    voices = engine.getProperty("voices")
    engine.setProperty("voice", voices[1].id)
    engine.setProperty("rate", 150)
    engine.say(words)
    engine.runAndWait()

def get_wechat_pid():
    for i in process_iter():
        pid_dic = i.as_dict(attrs=["pid", "name"])
        if pid_dic["name"] == "WeChat.exe":
            return pid_dic["pid"]

def log(text, level="normal", *, logfile_only=False):
    if level == "normal":
        print(Fore.LIGHTGREEN_EX + text)
    elif level == "info":
        if not logfile_only:
            print(Fore.LIGHTGREEN_EX + "[%s] [%s]: %s" % (
                time.strftime(r"%Y-%m-%d %H:%M:%S"), "INFO", text))
        logging.info(text)
    elif level == "warning":
        if not logfile_only:
            print(Fore.LIGHTYELLOW_EX + "[%s] [%s]: %s" % (
                time.strftime(r"%Y-%m-%d %H:%M:%S"), "WARNING", text))
        logging.warning(text)
        print(Fore.LIGHTGREEN_EX, end="")
    elif level == "error":
        if not logfile_only:
            print(Fore.LIGHTRED_EX + "[%s] [%s]: %s" % (
                time.strftime(r"%Y-%m-%d %H:%M:%S"), "ERROR", text))
        logging.error(text)
        print(Fore.LIGHTGREEN_EX, end="")
    elif level == "exception":
        if not logfile_only:
            print(Fore.LIGHTRED_EX + "[%s] [%s]: %s" % (
                time.strftime(r"%Y-%m-%d %H:%M:%S"), "ERROR", text))
        logging.exception(text)
        print(Fore.LIGHTGREEN_EX, end="")
    elif level == "debug":
        if config.debug:
            if not logfile_only:
                print(Fore.LIGHTBLUE_EX + "[%s] [%s]: %s" % (
                    time.strftime(r"%Y-%m-%d %H:%M:%S"), "DEBUG", text))
        logging.debug(text)
        print(Fore.LIGHTGREEN_EX, end="")
    else:
        if not logfile_only:
            print("[%s] [%s]: %s" % (
                time.strftime(r"%Y-%m-%d %H:%M:%S"), level, text))
        logging.info(text)
        print(Fore.LIGHTGREEN_EX, end="")

def wechat(text, executant_wrapper_object, *, with_spaces=True):
    pyautogui.hotkey("ctrl", "alt", "w")
    executant_wrapper_object.click_input()
    time.sleep(0.1)
    executant_wrapper_object.type_keys("[FTF] %s"%text, with_spaces=with_spaces)
    pyautogui.hotkey("enter")

def choice(choose="YN", text="Y/N", default=None, timeout=10, *, hide=False):
    choice = "choice /C "
    if " " in text:
        raise SyntaxError(
            "显示的文字中不能含有空格"
        )
    if hide:
        choice = choice + choose + " /N " + " /M " + text
    else:
        choice = choice + choose + " /M " + text
    if default:
        if default in choose:
            choice = choice + " /D " + default
            choice = choice + " /T " + str(timeout)
        else:
            raise SyntaxError(
                "按键默认值不在设置的按键中"
            )
    return os.system(choice)

def copyfile(*files):
    file_get_item = ""
    for i in files:
        if not os.path.exists(i):
            continue
        file_get_item += i + ","
    file_get_item = file_get_item[0:len(file_get_item) - 1]
    args = ["powershell", "Get-Item %s | Set-Clipboard"%file_get_item]
    subprocess.Popen(args)

log("初始化终端", "info")
