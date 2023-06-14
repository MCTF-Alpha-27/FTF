import os
import glob
import importlib
from cmd import Cmd
from typing import IO
from time import sleep
from .exceptions import *
from .functions import log

def loadcmd():
    for i in glob.glob("libs\\ExternalCommands\\cmd_*.py"):
        plugin_file_name = i.replace("\\", ".")[0:-3]
        log("External Commands: " + plugin_file_name, "info")
        importlib.import_module(plugin_file_name)

class FTFCmd(Cmd):
    intro = "欢迎使用《朝花夕拾协议》命令行\n"
    prompt = "[FTF Command Line] "
    doc_header = "help命令找到了以下命令的帮助文档，输入help <命令>来查看其帮助文档:"
    undoc_header = "help命令没有找到以下命令的帮助文档，也许这些命令有显示自己帮助文档的参数，试试输入<命令> /?来查看它们:"
    misc_header = "其它帮助命令:"
    nohelp = "没有找到%s命令的帮助文档，也许这个命令有显示自己帮助文档的参数，试试输入<命令> /?来查看它们"

    def __init__(self, completekey: str = "tab", stdin: IO[str] | None = None, stdout: IO[str] | None = None) -> None:
        super().__init__(completekey, stdin, stdout)

    def emptyline(self):
        return
    
    def preloop(self) -> None:
        os.system("cls")

    def default(self, line: str) -> None:
        os.system(line)

    def onecmd(self, line: str) -> bool:
        if not line == "" and not line.isspace():
            log(self.prompt + line, "info")
        return super().onecmd(line)

    def do_exit(self, args: str):
        """退出《朝花夕拾协议》命令行"""
        if args.split(" ")[0] == "/?":
            print(self.do_exit.__doc__)
            return
        loadcmd()
        raise CommandLineExit()
    
def help_ftf():
    for i in range(2):
        os.system("color 0c")
        sleep(0.01)
        os.system("color 0a")
        sleep(0.01)
    sleep(2)
    os.system("color 0c")
    raise AdminMode()

class FTFAdminCmd(FTFCmd):
    intro = "欢迎您，协议创始人\n一切为了不远后的旧事重提\n"
    prompt = "[FTF ADMIN] "

    def onecmd(self, line: str) -> bool:
        result = super().onecmd(line)
        os.system("color 0c")
        return result

def help_ftf_admin():
    print("您已处在协议创始人权限下")

loadcmd()
