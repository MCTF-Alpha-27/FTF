"""
配置文件解析
"""
from configparser import ConfigParser
import os

config = ConfigParser()
if not os.path.exists("config.ini"):
    config["Logger"] = {"debug": False}
    config["Docs"] = {"client_id": r"{client_id}", "access_token": r"{access_token}", "open_id": r"{open_id}"}
    config["Feishu"] = {"app_id": r"{app_id}", "app_secret": r"{app_secret}", "allowed_open_ids": r"{allowed_open_ids}"}
    with open("config.ini", "w", encoding="utf-8") as cfgfile:
        config.write(cfgfile)

config.read("config.ini", encoding="utf-8")
debug = True if config.get("Logger", "debug") == "True" else False
ftfpath = config.get("FTF", "ftfpath")
client_id = config.get("Docs", "client_id")
access_token = config.get("Docs", "access_token")
open_id = config.get("Docs", "open_id")

def _option(section, option, default=""):
    if config.has_option(section, option):
        return config.get(section, option).strip()
    return default

feishu_app_id = _option("Feishu", "app_id")
feishu_app_secret = _option("Feishu", "app_secret")
feishu_encrypt_key = _option("Feishu", "encrypt_key")
feishu_verification_token = _option("Feishu", "verification_token")
feishu_allowed_open_ids = [
    i.strip()
    for i in _option("Feishu", "allowed_open_ids").replace("，", ",").split(",")
    if i.strip()
]
