import os
from configparser import ConfigParser

config = ConfigParser()
if not os.path.exists("config.ini"):
    config["Logger"] = {"debug": False, "voice": True}
    config["FTF"] = {"ftfpath": "{ftfpath}", "controller": "一只叫迷迭香的菲林"}
    with open("config.ini", "w", encoding="gbk") as cfgfile:
        config.write(cfgfile)

config.read("config.ini")
debug = True if config.get("Logger", "debug") == "True" else False
voice = True if config.get("Logger", "voice") == "True" else False
ftfpath = config.get("FTF", "ftfpath")
controller = config.get("FTF", "controller")
