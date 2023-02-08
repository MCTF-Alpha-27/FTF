import pyttsx3
import time
import logging
import os
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
        print(text)
    elif level == "info":
        print("[%s] [%s]: %s" % (
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
    else:
        print("[%s] [%s]: %s" % (
            time.strftime(r"%Y-%m-%d %H:%M:%S"), level, text))
        logging.info(text)
