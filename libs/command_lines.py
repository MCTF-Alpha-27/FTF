import os
import re
import importlib
from cmd import Cmd
from glob import glob
from typing import IO
from time import sleep
from docx import Document
from .exceptions import *
from .functions import log, choice
from .config import ftfpath

def loadcmd():
    for i in glob("libs\\ExternalCommands\\cmd_*.py"):
        plugin_file_name = i.replace("\\", ".")[0:-3]
        log("External Commands: " + plugin_file_name, "debug")
        importlib.import_module(plugin_file_name)

class FTFCmd(Cmd):
    intro = "欢迎使用《朝花夕拾协议》命令行\n"
    prompt = "[FTF Command Line] "
    doc_header = "help命令找到了以下命令的帮助文档，输入help <命令>来查看其帮助文档:"
    undoc_header = "help命令没有找到以下命令的帮助文档，也许这些命令有显示自己帮助文档的参数，试试输入<命令> /?来查看它们:"
    misc_header = "其它帮助命令:"
    nohelp = "没有找到%s命令的帮助文档，也许这个命令有显示自己帮助文档的参数，试试输入<命令> /?来查看它们"

    def __init__(self, completekey: str = "tab", stdin: IO[str] | None = None, stdout: IO[str] | None = None) -> None:
        self.years = [i for i in os.listdir(ftfpath) if os.path.isdir(os.path.join(ftfpath, i))]
        super().__init__(completekey, stdin, stdout)

    def emptyline(self):
        return
    
    def preloop(self) -> None:
        os.system("cls")
        loadcmd()

    def default(self, line: str) -> None:
        os.system(line)

    def onecmd(self, line: str) -> bool:
        if not line == "" and not line.isspace():
            log(self.prompt + line, "info", logfile_only=True)
        return super().onecmd(line)

    def do_exit(self, args: str):
        """
        退出《朝花夕拾协议》命令行。

        语法：exit [/?]
            无参数  退出《朝花夕拾协议》命令行。
            /?      显示此帮助文档。
        """
        if args.split(" ")[0] == "/?":
            print(self.do_exit.__doc__)
            return
        raise CommandLineExit()
    
    def do_find(self, args: str):
        """
        查找事件记录文档中的关键字词。

        语法：find <keywords> in [year <years> | month <months> | *] [/?]
            keywords        事件记录文档中的关键字词。
            year <years>    事件记录文档所在的年份，可填多个。
            month <months>  事件记录文档所在的月份，可填多个。
            *               在所有文档中查找关键字词。
            /?              显示此帮助文档。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_find.__doc__)
            return
        keywords = args.split(" in ")[0].split(" ")
        documents = args.split(" in ")[1].split(" ")
        if documents[0] == "*":
            count = 0
            for year in self.years:
                docments_path = os.path.join(ftfpath, year)
                for document in glob(f"{docments_path}\\*.docx"):
                    doc = Document(document)
                    line = 0
                    for paragraph in doc.paragraphs:
                        line += 1
                        for keyword in keywords:
                            if keyword in paragraph.text:
                                print(f"在{document}第{line}个段落中找到关键字词: {keyword}")
                                log(f"在{document}第{line}个段落中找到关键字词: {keyword}", "info", logfile_only=True)
                                print(f"    -> {paragraph.text}")
                                log(f"  -> {paragraph.text}", "info", logfile_only=True)
                                print()
                                log("", "info", logfile_only=True)
                                count += 1
            print(f"在所有文档中共发现{count}个关键字词\n")
            log(f"在所有文档中共发现{count}个关键字词", "info", logfile_only=True)
            log("", "info", logfile_only=True)
        elif documents[0] == "year":
            year = documents[1:]
            count = 0
            for i in year:
                if i not in self.years:
                    print(f"未找到{i}年的事件记录文档")
                    log(f"未找到{i}年的事件记录文档", "warning", logfile_only=True)
                    year.remove(i)
                    continue
                docments_path = os.path.join(ftfpath, i)
                for document in glob(f"{docments_path}\\*.docx"):
                    doc = Document(document)
                    line = 0
                    for paragraph in doc.paragraphs:
                        line += 1
                        for keyword in keywords:
                            if keyword in paragraph.text:
                                print(f"在{document}第{line}个段落中找到关键字词: {keyword}")
                                log(f"在{document}第{line}个段落中找到关键字词: {keyword}", "info", logfile_only=True)
                                print(f"    -> {paragraph.text}")
                                log(f"  -> {paragraph.text}", "info", logfile_only=True)
                                print()
                                log("", "info", logfile_only=True)
                                count += 1
            print(f"在{", ".join(year)}这{len(year)}年的事件记录文档中共发现{count}个关键字词\n")
            log(f"在{", ".join(year)}这{len(year)}年的事件记录文档中共发现{count}个关键字词", "info", logfile_only=True)
            log("", "info", logfile_only=True)
        elif documents[0] == "month":
            month = documents[1:]
            count = 0
            for i in month:
                if i not in [str(i) for i in range(1, 13)]:
                    print(f"未找到{i}月的事件记录文档")
                    log(f"未找到{i}月的事件记录文档", "warning", logfile_only=True)
                    month.remove(i)
                    continue
                for j in self.years:
                    docments_path = document = os.path.join(ftfpath, j, f"{i}月.docx")
                    if not os.path.exists(docments_path):
                        print(f"未找到{j}年{i}月的事件记录文档")
                        log(f"未找到{j}年{i}月的事件记录文档", "warning", logfile_only=True)
                        continue
                    doc = Document(docments_path)
                    line = 0
                    for paragraph in doc.paragraphs:
                        line += 1
                        for keyword in keywords:
                            if keyword in paragraph.text:
                                print(f"在{document}第{line}个段落中找到关键字词: {keyword}")
                                log(f"在{document}第{line}个段落中找到关键字词: {keyword}", "info", logfile_only=True)
                                print(f"    -> {paragraph.text}")
                                log(f"  -> {paragraph.text}", "info", logfile_only=True)
                                print()
                                log("", "info", logfile_only=True)
                                count += 1
            print(f"在{", ".join(month)}月这{len(month)}个月的事件记录文档中共发现{count}个关键字词\n")
            log(f"在{", ".join(month)}月这{len(month)}个月的事件记录文档中共发现{count}个关键字词", "info", logfile_only=True)
            log("", "info", logfile_only=True)
        else:
            print(self.do_find.__doc__)

    def complete_find(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        #if line.endswith("year "):
            #return self.years
        #if line.endswith("month "):
            #return [str(i) for i in range(1, 13)]
        if re.match(r"find \w* in year \w*", line):
            return [i for i in self.years if i.startswith(text)]
        if re.match(r"find \w* in month \w*", line):
            return [i for i in [str(i) for i in range(1, 13)] if i.startswith(text)]
        if re.match(r"find \w* in ", line):
            return [i for i in ["*", "year", "month"] if i.startswith(text)]
        if len(line.split(" ")) > 2:
            return ["in"]
        return []

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
        if line == "" or line.isspace():
            return
        c = choice("YN", "协议创始人，请牢记您的命令一旦执行便无法撤回，您确定要执行此命令吗")
        if c == 2:
            return
        result = super().onecmd(line)
        os.system("color 0c")
        return result

def help_ftf_admin():
    print("您已处在协议创始人权限下")
