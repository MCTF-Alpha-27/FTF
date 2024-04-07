import os
from configparser import ConfigParser

config = ConfigParser()
if not os.path.exists("config.ini"):
    config["FTF"] = {"debug": False}
    with open("config.ini", "w") as cfgfile:
        config.write(cfgfile)

config.read("config.ini")
debug = bool(config.get("FTF", "debug"))
