"""
《朝花夕拾协议》监听终端（飞书版）
"""
from . import config
from .feishu import FeishuTransport
from .functions import log

def main() -> int:
    if not config.feishu_app_id or config.feishu_app_id.startswith("{"):
        log("未配置飞书 App ID，请在 config.ini 的 [Feishu] 段填写 app_id", "error")
        return 1
    if not config.feishu_app_secret or config.feishu_app_secret.startswith("{"):
        log("未配置飞书 App Secret，请在 config.ini 的 [Feishu] 段填写 app_secret", "error")
        return 1

    log("监听终端启动中", "info")
    if config.feishu_allowed_open_ids:
        log(f"已配置指令白名单，共{len(config.feishu_allowed_open_ids)}个 open_id", "info")
    else:
        log("未配置指令白名单，任何能给机器人发消息的人都可以执行指令", "warning")
        log("可先发送 /whoami 查看自己的 open_id，再填入 config.ini 的 allowed_open_ids", "warning")
    log("正在连接到飞书", "info")

    transport = FeishuTransport(
        config.feishu_app_id,
        config.feishu_app_secret,
        config.feishu_allowed_open_ids,
        config.feishu_encrypt_key,
        config.feishu_verification_token,
    )
    try:
        result = transport.run()
    except KeyboardInterrupt:
        log("监听终端已停止", "info")
        return 0
    except Exception as e:
        log(f"监听终端异常退出: {e}", "error")
        log("详细信息已写入日志", "exception")
        return 1
    return 1 if result == "error" else 0
