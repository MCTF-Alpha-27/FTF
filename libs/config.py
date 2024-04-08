import os
from configparser import ConfigParser

config = ConfigParser()
if not os.path.exists("config.ini"):
    config["Logger"] = {"debug": False}
    config["FTF"] = {"ftfpath": "{ftfpath}"}
    with open("config.ini", "w", encoding="gbk") as cfgfile:
        config.write(cfgfile)

config.read("config.ini")
debug = bool(config.get("Logger", "debug"))
ftfpath = config.get("FTF", "ftfpath")
