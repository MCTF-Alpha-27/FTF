import os
import re
import importlib
from cmd import Cmd
from glob import glob
from typing import IO
from time import sleep
from docx import Document
from colorama import Fore, init
from collections import OrderedDict
from .exceptions import *
from .functions import log, choice
from .config import ftfpath

init()

def loadcmd():
    for i in glob("libs\\ExternalCommands\\cmd_*.py"):
        plugin_file_name = i.replace("\\", ".")[0:-3]
        log("External Commands: " + plugin_file_name, "debug")
        importlib.import_module(plugin_file_name)

class FTFCmd(Cmd):
    intro = "欢迎使用《朝花夕拾协议》命令行\n"
    prompt = "[FTF Command Line] "
    doc_header = "发现了以下命令的帮助文档，输入help <命令>来查看其帮助文档:"
    undoc_header = "未发现以下命令的帮助文档，也许这些命令有显示自己帮助文档的参数，试试输入<命令> /?来查看它们:"
    misc_header = "其它帮助命令:"
    nohelp = "没有找到%s命令的帮助文档，也许这个命令有显示自己帮助文档的参数，试试输入<命令> /?来查看它们"

    def __init__(self, completekey: str = "tab", stdin: IO[str] | None = None, stdout: IO[str] | None = None) -> None:
        super().__init__(completekey, stdin, stdout)
        self.years = [i for i in os.listdir(ftfpath) if os.path.isdir(os.path.join(ftfpath, i))]

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

        语法：find <keywords>/<regex> in [year <years> | month <months> | *] [/?]
            keywords        事件记录文档中的关键字词。
                            使用空格分隔多个关键字词，表示查找包含多个关键字词的项。
                            使用“&”符号连接多个关键字词，表示查找的项中必须同时包含这些关键字词。
            regex           正则表达式。以“/”开头并以“/search”或“/match”结尾的字符串将被视为正则表达式。
                            正则表达式以“/search”结尾表示从整个字符串中搜索匹配项。
                            正则表达式以“/match”结尾表示从字符串开头匹配。
            year <years>    事件记录文档所在的年份，可填多个，使用空格分隔。
            month <months>  事件记录文档所在的月份，可填多个，使用空格分隔。
            *               在所有文档中查找关键字词。
            /?              显示此帮助文档。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_find.__doc__)
            return
        keywords = set(args.split(" in ")[0].split(" "))
        documents = list(OrderedDict.fromkeys(args.split(" in ")[1].split(" ")).keys())
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
                            if keyword.startswith("/") and keyword.endswith("/search"):
                                if re.search(keyword[1:-8], paragraph.text):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}项（从整个字符串中搜索匹配项） -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            elif keyword.startswith("/") and keyword.endswith("/match"):
                                if re.match(keyword[1:-6], paragraph.text):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            elif "&" in keyword:
                                if all(i in paragraph.text for i in keyword.split("&")):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{",".join(keyword.split("&"))}”的项 -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{",".join(keyword.split("&"))}”的项 -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            else:
                                if keyword in paragraph.text:
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}", "info", logfile_only=True)
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
                            if keyword.startswith("/") and keyword.endswith("/search"):
                                if re.search(keyword[1:-8], paragraph.text):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}项（从整个字符串中搜索匹配项） -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            elif keyword.startswith("/") and keyword.endswith("/match"):
                                if re.match(keyword[1:-6], paragraph.text):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            elif "&" in keyword:
                                if all(i in paragraph.text for i in keyword.split("&")):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{",".join(keyword.split("&"))}”的项 -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{",".join(keyword.split("&"))}”的项 -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            else:
                                if keyword in paragraph.text:
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}", "info", logfile_only=True)
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
                            if keyword.startswith("/") and keyword.endswith("/search"):
                                if re.search(keyword[1:-8], paragraph.text):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}项（从整个字符串中搜索匹配项） -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            elif keyword.startswith("/") and keyword.endswith("/match"):
                                if re.match(keyword[1:-6], paragraph.text):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            elif "&" in keyword:
                                if all(i in paragraph.text for i in keyword.split("&")):
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{",".join(keyword.split("&"))}”的项 -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{",".join(keyword.split("&"))}”的项 -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
                            else:
                                if keyword in paragraph.text:
                                    print(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}")
                                    log(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}", "info", logfile_only=True)
                                    count += 1
            print(f"在{", ".join(month)}月这{len(month)}个月的事件记录文档中共发现{count}个关键字词\n")
            log(f"在{", ".join(month)}月这{len(month)}个月的事件记录文档中共发现{count}个关键字词", "info", logfile_only=True)
            log("", "info", logfile_only=True)
        else:
            print(self.do_find.__doc__)

    def complete_find(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        if re.match(r'find [\w!@#$%^&*(),.?":{}|<>/]+ in year \w*', line):
            return [i for i in self.years if i.startswith(text)]
        if re.match(r'find [\w!@#$%^&*(),.?":{}|<>/]+ in month \w*', line):
            return [i for i in [str(i) for i in range(1, 13)] if i.startswith(text)]
        if re.match(r'find [\w!@#$%^&*(),.?":{}|<>/]+ in ', line):
            return [i for i in ["*", "year", "month"] if i.startswith(text)]
        if len(line.split(" ")) > 2:
            return ["in"]
        return []

def help_ftf():
    for i in range(2):
        os.system("color 0c")
        sleep(0.1)
        os.system("color 0a")
        sleep(0.1)
    sleep(2)
    raise AdminMode()

class FTFAdminCmd(FTFCmd):
    intro = Fore.LIGHTRED_EX + "欢迎您，协议创始人\n一切为了不远后的旧事重提\n"
    prompt = "[FTF ADMIN] "

    def onecmd(self, line: str) -> bool:
        if line == "" or line.isspace():
            return
        c = choice("YN", Fore.LIGHTRED_EX + "协议创始人，请牢记您的命令一旦执行便无法撤回，您确定要执行此命令吗")
        os.system("color 0c")
        if c == 2:
            return
        return super().onecmd(line)

def help_ftf_admin():
    print(Fore.LIGHTRED_EX + "您已处在协议创始人权限下")
