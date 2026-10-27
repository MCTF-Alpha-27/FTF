"""
《朝花夕拾协议》异常类
"""

class TransferTerminalControl(Exception):
    """传输终端控制异常。

    历史遗留类，当前代码未使用，保留以兼容外部调用。
    """

class CommandLineExit(Exception):
    """请求退出《朝花夕拾协议》命令行。"""

class AdminMode(Exception):
    """在公开命令行中请求切换到创始人模式。"""

class ConfigError(Exception):
    """配置文件缺失或配置项无效。"""

class ConfirmationRequired(Exception):
    """命令执行到需要人工确认的步骤。

    question 保存确认提示，由调用方决定确认或取消。
    """

    def __init__(self, question: str) -> None:
        super().__init__(question)
        self.question = question

class AIError(Exception):
    """AI 服务调用或响应解析失败。

    异常消息可以直接展示给用户，不包含 API 密钥。
    """
