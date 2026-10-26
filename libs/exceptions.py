"""
异常类
"""
class TransferTerminalControl(Exception):
    pass

class CommandLineExit(Exception):
    pass

class AdminMode(Exception):
    pass

class ConfigError(Exception):
    pass
