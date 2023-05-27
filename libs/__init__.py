import pyttsx3
import time
import logging
import os
import pyautogui
from psutil import process_iter
from colorama import Fore, init

if not os.path.exists("logs"):
    os.mkdir("logs")

engine = pyttsx3.init()
logging.basicConfig(filename="logs/%s.log"%time.strftime(r"%Y-%m-%d-%H.%M.%S"), level=logging.INFO, format="[%(asctime)s] [%(levelname)s]: %(message)s")

def say_in_english(words):
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

logfile_ = True
def log(text, level="normal"):
    init()
    if level == "normal":
        print(Fore.GREEN + text)
    elif level == "info":
        print(Fore.GREEN + "[%s] [%s]: %s" % (
            time.strftime(r"%Y-%m-%d %H:%M:%S"), "INFO", text))
        logging.info(text)
    elif level == "warning":
        print(Fore.YELLOW + "[%s] [%s]: %s" % (
            time.strftime(r"%Y-%m-%d %H:%M:%S"), "WARNING", text))
        logging.warning(text)
    elif level == "error":
        print(Fore.RED + "[%s] [%s]: %s" % (
            time.strftime(r"%Y-%m-%d %H:%M:%S"), "ERROR", text))
        logging.error(text)
    elif level == "debug":
        print(Fore.WHITE + "[%s] [%s]: %s" % (
            time.strftime(r"%Y-%m-%d %H:%M:%S"), "DEBUG", text))
        logging.debug(text)
    else:
        print("[%s] [%s]: %s" % (
            time.strftime(r"%Y-%m-%d %H:%M:%S"), level, text))
        logging.info(text)

def wechat(text, executant_wrapper_object, *, with_spaces=True):
    pyautogui.hotkey("ctrl", "alt", "w")
    executant_wrapper_object.click_input()
    executant_wrapper_object.type_keys("[FTF] %s"%text, with_spaces=with_spaces)
    pyautogui.hotkey("enter")
