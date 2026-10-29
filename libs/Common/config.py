"""
配置文件解析
"""

from configparser import ConfigParser
import os

config = ConfigParser()
if not os.path.exists("config.ini"):
    config["Logger"] = {"debug": False}
    config["FTF"] = {"ftfpath": r"{ftfpath}"}
    config["Docs"] = {
        "client_id": r"{client_id}",
        "access_token": r"{access_token}",
        "open_id": r"{open_id}",
    }
    config["Feishu"] = {
        "app_id": r"{app_id}",
        "app_secret": r"{app_secret}",
        "allowed_open_ids": r"{allowed_open_ids}",
    }
    config["AI"] = {
        "enabled": False,
        "api_key": r"{api_key}",
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen3-8B",
        "timeout": 60,
        "max_tokens": 1024,
        "temperature": 0.2,
        "enable_thinking": False,
        "protocol_in_prompt": True,
    }
    with open("config.ini", "w", encoding="utf-8") as cfgfile:
        config.write(cfgfile)

config.read("config.ini", encoding="utf-8")


def _option(section, option, default=""):
    """读取字符串配置，缺失或为空时回退默认值"""
    if config.has_option(section, option):
        return config.get(section, option).strip() or default
    return default


def _number(section, option, default, cast):
    """读取数字配置，缺失或格式错误时回退默认值"""
    value = _option(section, option)
    if not value:
        return default
    try:
        return cast(value)
    except ValueError:
        return default


def _flag(section, option, default=False):
    """读取布尔配置，接受 True/true/1/yes 等写法"""
    value = _option(section, option)
    if not value:
        return default
    return value.lower() in ("1", "true", "yes", "on")


debug = _flag("Logger", "debug")
ftfpath = _option("FTF", "ftfpath", r"{ftfpath}")
client_id = _option("Docs", "client_id", r"{client_id}")
access_token = _option("Docs", "access_token", r"{access_token}")
open_id = _option("Docs", "open_id", r"{open_id}")

feishu_app_id = _option("Feishu", "app_id")
feishu_app_secret = _option("Feishu", "app_secret")
feishu_encrypt_key = _option("Feishu", "encrypt_key")
feishu_verification_token = _option("Feishu", "verification_token")
feishu_allowed_open_ids = [
    i.strip()
    for i in _option("Feishu", "allowed_open_ids")
    .replace("，", ",")
    .split(",")
    if i.strip()
]

ai_enabled = _flag("AI", "enabled")
ai_api_key = _option("AI", "api_key")
ai_base_url = _option("AI", "base_url", "https://api.siliconflow.cn/v1")
ai_model = _option("AI", "model", "Qwen/Qwen3-8B")
ai_timeout = _number("AI", "timeout", 60.0, float)
ai_max_tokens = _number("AI", "max_tokens", 1024, int)
ai_temperature = _number("AI", "temperature", 0.2, float)
ai_enable_thinking = _flag("AI", "enable_thinking")
ai_protocol_in_prompt = _flag("AI", "protocol_in_prompt", True)
