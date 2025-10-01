from cmd import Cmd
from glob import glob
from typing import IO
from docx import Document
from time import sleep
from colorama import Fore, init
from collections import OrderedDict
from .exceptions import *
from .functions import log, choice
from .config import ftfpath

import os
import re
import importlib
import datetime
import math

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
        try:
            callback = super().onecmd(line)
        except Exception as e:
            if isinstance(e, AdminMode) or isinstance(e, CommandLineExit):
                raise e
            print(f"{Fore.LIGHTRED_EX}{e.__class__.__name__}: {str(e)}")
            log(f"{e.__class__.__name__}: {str(e)}", "error", logfile_only=True)
            print(f"{Fore.LIGHTRED_EX}运行时发生错误，详细信息已被写入日志")
            log("运行时发生错误，详细信息已被写入日志", "exception", logfile_only=True)
            return False
        return callback

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
    
    def _find(self, document: str, keywords: set[str], count: int) -> int:
        doc = Document(document)
        line = 0
        month_span = None
        month_span_count = 0
        month_span_count_for_L = 0
        space_count = 0
        space_count_now = 0
        for paragraph in doc.paragraphs:
            if paragraph.text == "跨月":
                month_span = math.inf
                month_span_count += 1
            if paragraph.text == "":
                space_count += 1
            if "月度评估" in paragraph.text:
                space_count -= 1
        for paragraph in doc.paragraphs:
            line += 1
            if paragraph.text == "跨月":
                month_span = line
                month_span_count_for_L += 1
            if paragraph.text == "" and month_span_count > 0:
                space_count_now += 1
                if month_span_count == 1:
                    month_span = None
                elif month_span_count == 2:
                    if space_count_now == space_count:
                        month_span = math.inf
                    else:
                        month_span = None
                else:
                    print(f"在{document}中发现两个以上的跨月事件，请检查文档")
                    log(f"在{document}中发现两个以上的跨月事件，请检查文档", "warning", logfile_only=True)
                    return count
            for keyword in keywords:
                if keyword.startswith("/") and keyword.endswith("/search"):
                    if re.search(keyword[1:-7], paragraph.text):
                        if month_span and keyword != "跨月":
                            if line > month_span:
                                print(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}")
                                log(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}", "info", logfile_only=True)
                            else:
                                if space_count_now < space_count:
                                    month_span_count_for_E = 1
                                else:
                                    month_span_count_for_E = 2
                                print(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}")
                                log(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}", "info", logfile_only=True)
                        else:
                            print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}")
                            log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-6]}的项（从整个字符串中搜索匹配项） -> {paragraph.text}", "info", logfile_only=True)
                        count += 1
                elif keyword.startswith("/") and keyword.endswith("/match"):
                    if re.match(keyword[1:-6], paragraph.text):
                        if month_span and keyword != "跨月":
                            if line > month_span:
                                print(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}")
                                log(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}", "info", logfile_only=True)
                            else:
                                if space_count_now < space_count:
                                    month_span_count_for_E = 1
                                else:
                                    month_span_count_for_E = 2
                                print(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}")
                                log(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}", "info", logfile_only=True)
                        else:
                            print(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}")
                            log(f"{count + 1}. 在{document}第{line}个段落中找到符合正则表达式{keyword[0:-5]}的项（从字符串开头匹配） -> {paragraph.text}", "info", logfile_only=True)
                        count += 1
                elif "&" in keyword:
                    if all(i in paragraph.text for i in keyword.split("&")):
                        if month_span and keyword != "跨月":
                            if line > month_span:
                                print(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到同时包含关键字词“{','.join(keyword.split('&'))}”的项 -> {paragraph.text}")
                                log(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到同时包含关键字词“{','.join(keyword.split('&'))}”的项 -> {paragraph.text}", "info", logfile_only=True)
                            else:
                                if space_count_now < space_count:
                                    month_span_count_for_E = 1
                                else:
                                    month_span_count_for_E = 2
                                print(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到同时包含关键字词“{','.join(keyword.split('&'))}”的项 -> {paragraph.text}")
                                log(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到同时包含关键字词“{','.join(keyword.split('&'))}”的项 -> {paragraph.text}", "info", logfile_only=True)
                        else:
                            print(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{','.join(keyword.split('&'))}”的项 -> {paragraph.text}")
                            log(f"{count + 1}. 在{document}第{line}个段落中找到同时包含关键字词“{','.join(keyword.split('&'))}”的项 -> {paragraph.text}", "info", logfile_only=True)
                        count += 1
                else:
                    if keyword in paragraph.text:
                        if month_span and keyword != "跨月":
                            if line > month_span:
                                print(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}")
                                log(f"{count + 1}. [L{month_span_count_for_L}]在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}", "info", logfile_only=True)
                            else:
                                if space_count_now < space_count:
                                    month_span_count_for_E = 1
                                else:
                                    month_span_count_for_E = 2
                                print(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}")
                                log(f"{count + 1}. [E{month_span_count_for_E}]在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}", "info", logfile_only=True)
                        else:
                            print(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}")
                            log(f"{count + 1}. 在{document}第{line}个段落中找到关键字词: {keyword} -> {paragraph.text}", "info", logfile_only=True)
                        count += 1
        return count

    def do_find(self, args: str):
        """
        查找文档中的关键字词。

        语法：find <keywords>/<regex> in [year <years> | month <months> | only <document> | *] [/?]
            keywords        指定查找的关键字词。
                            使用空格分隔多个关键字词，表示查找包含多个关键字词的项。
                            使用“&”符号连接多个关键字词，表示查找的项中必须同时包含这些关键字词。
            regex           正则表达式。以“/”符号开头并以“/search”或“/match”结尾的字符串将被视为正则表达式。
                            正则表达式以“/search”结尾表示从整个字符串中搜索匹配项。
                            正则表达式以“/match”结尾表示从字符串开头匹配。
            year <years>    文档所属的年份，可填多个，使用空格分隔。
            month <months>  文档所属的月份，可填多个，使用空格分隔。
            only <document> 仅在指定的文档中查找关键字词。
                            若指定的文档格式为<year>/<month>，则表示在事件记录文档中查找，如“2023/1”表示2023年1月的事件记录文档。
                            若指定的文档格式为<year>/annual_summary，则表示在年度总结中查找，如2023/annual_summary表示2023年的年度总结。
            *               在所有文档中查找关键字词，包括年度总结。
            /?              显示此帮助文档。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_find.__doc__)
            return
        keywords = set(args.split(" in ")[0].split(" "))
        documents = list(OrderedDict.fromkeys(args.split(" in ")[1].split(" ")).keys())
        print("注意：查询涉及的文档中若存在跨月事件，可能导致查询事件所属月份不准确，请注意核实")
        log("注意：查询涉及的文档中若存在跨月事件，可能导致查询事件所属月份不准确，请注意核实", "info", logfile_only=True)
        print("[E1/E2]表示该项发生在第一个/第二个跨月事件之前，[L1/L2]表示该项发生在第一个/第二个跨月事件之后，以文档中每个周度周期内各自的跨月事件为相对位置")
        log("[E1/E2]表示该项发生在第一个/第二个跨月事件之前，[L1/L2]表示该项发生在第一个/第二个跨月事件之后，以文档中每个周度周期内各自的跨月事件为相对位置", "info", logfile_only=True)
        print("文档内最多存在两个跨月事件，若存在两个以上的跨月事件，视为记录错误")
        if documents[0] == "*":
            count = 0
            for year in self.years:
                docments_path = os.path.join(ftfpath, year)
                for document in glob(f"{docments_path}\\*.docx"):
                    count = self._find(document, keywords, count)
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
                    count = self._find(document, keywords, count)
            print(f"在{', '.join(year)}这{len(year)}年的事件记录文档中共发现{count}个关键字词\n")
            log(f"在{', '.join(year)}这{len(year)}年的事件记录文档中共发现{count}个关键字词", "info", logfile_only=True)
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
                    docments_path = os.path.join(ftfpath, j, f"{i}月.docx")
                    if not os.path.exists(docments_path):
                        print(f"未找到{j}年{i}月的事件记录文档")
                        log(f"未找到{j}年{i}月的事件记录文档", "warning", logfile_only=True)
                        continue
                    count = self._find(docments_path, keywords, count)
            print(f"在{', '.join(month)}月这{len(month)}个月的事件记录文档中共发现{count}个关键字词\n")
            log(f"在{', '.join(month)}月这{len(month)}个月的事件记录文档中共发现{count}个关键字词", "info", logfile_only=True)
            log("", "info", logfile_only=True)
        elif documents[0] == "only":
            count = 0
            document = documents[1]
            if document.split("/")[1] == "annual_summary":
                docments_path = os.path.join(ftfpath, document.split("/")[0], "年度总结.docx")
            else:
                docments_path = os.path.join(ftfpath, document.split("/")[0], f"{document.split('/')[1]}月.docx")
            if not os.path.exists(docments_path):
                if document.split("/")[1] == "annual_summary":
                    print(f"未找到{document.split('/')[0]}年的年度总结")
                    log(f"未找到{document.split('/')[0]}年的年度总结", "warning", logfile_only=True)
                else:
                    print(f"未找到{document}的事件记录文档")
                    log(f"未找到{document}的事件记录文档", "warning", logfile_only=True)
                return
            count = self._find(docments_path, keywords, count)
            if document.split("/")[1] == "annual_summary":
                print(f"在{document.split('/')[0]}年的年度总结中共发现{count}个关键字词\n")
                log(f"在{document.split('/')[0]}年的年度总结中共发现{count}个关键字词", "info", logfile_only=True)
            else:
                print(f"在{document}这个事件记录文档中共发现{count}个关键字词\n")
                log(f"在{document}这个事件记录文档中共发现{count}个关键字词", "info", logfile_only=True)
            log("", "info", logfile_only=True)
        else:
            print(self.do_find.__doc__)

    def complete_find(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        if re.match(r"find [^.*$]+ in year \w*", line):
            return [i for i in self.years if i.startswith(text)]
        if re.match(r"find [^.*$]+ in month \w*", line):
            return [i for i in [str(i) for i in range(1, 13)] if i.startswith(text)]
        if re.match(r"find [^.*$]+ in only \w*", line):
            parts = text.split("/")
            year_part = parts[0] if len(parts) > 0 else ""
            month_part = parts[1] if len(parts) > 1 else ""
            if "/" not in text:
                return [f"{y}/" for y in self.years if y.startswith(year_part)]
            options = []
            for month in range(1, 13):
                month_str = str(month)
                if month_str.startswith(month_part):
                    options.append(f"{year_part}/{month_str}")
            if "annual_summary".startswith(month_part):
                options.append(f"{year_part}/annual_summary")
            return options
        if re.match(r"find [^.*$]+ in ", line):
            return [i for i in ["*", "year", "month", "only"] if i.startswith(text)]
        if len(line.split(" ")) > 2:
            return ["in"]
        return []
    
    def do_open(self, args: str):
        """
        打开指定的文档。

        语法：open <document> [/?]
            document    指定的文档。
                        若文档格式为<year>/<month>，则表示打开事件记录文档，如“2023/1”表示2023年1月的事件记录文档。
                        若文档格式为<year>/annual_summary，则表示打开年度总结，如2023/annual_summary表示2023年的年度总结。
            /?          显示此帮助文档。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_open.__doc__)
            return
        document = args
        if document.split("/")[1] == "annual_summary":
            docments_path = os.path.join(ftfpath, document.split("/")[0], "年度总结.docx")
        else:
            docments_path = os.path.join(ftfpath, document.split("/")[0], f"{document.split("/")[1]}月.docx")
        if not os.path.exists(docments_path):
            if document.split("/")[1] == "annual_summary":
                print(f"未找到{document.split('/')[0]}年的年度总结")
                log(f"未找到{document.split('/')[0]}年的年度总结", "warning", logfile_only=True)
            else:
                print(f"未找到{document}的事件记录文档")
                log(f"未找到{document}的事件记录文档", "warning", logfile_only=True)
            return
        os.system(f"start {docments_path}")
        if document.split("/")[1] == "annual_summary":
            log(f"打开了{document.split('/')[0]}年的年度总结", "info", logfile_only=True)
        else:
            log(f"打开了{document}的事件记录文档", "info", logfile_only=True)

    def complete_open(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        parts = text.split("/")
        year_part = parts[0] if len(parts) > 0 else ""
        month_part = parts[1] if len(parts) > 1 else ""
        if "/" not in text:
            return [f"{i}/" for i in self.years if i.startswith(year_part)]
        if month_part == "":
            months = [f"{year_part}/{i}" for i in [str(i) for i in range(1, 13)]]
            annual = [f"{year_part}/annual_summary"]
            return months + annual
        else:
            if month_part.startswith("a"):
                return [f"{year_part}/annual_summary"] if "annual_summary".startswith(month_part) else []
            else:
                return [f"{year_part}/{i}" for i in [str(i) for i in range(1, 13)] if i.startswith(month_part)]

    def do_count(self, args: str):
        """
        统计年度总结中的新事物（或重大事件）以及“时期”的数量。

        语法：count [N&M | normal_period | combined_period] in <year> [/?]
            N&M                 统计新事物（或重大事件）的数量。
            normal_period       统计常规“时期”的数量。
            combined_period     统计合称“时期”的数量。
            year                年度总结所属的年份。
            /?                  显示此帮助文档。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_count.__doc__)
            return
        if args.split(" ")[0] == "N&M":
            year = args.split(" ")[2]
            docments_path = os.path.join(ftfpath, year, "年度总结.docx")
            if not os.path.exists(docments_path):
                print(f"未找到{year}年的年度总结")
                log(f"未找到{year}年的年度总结", "warning", logfile_only=True)
                return
            doc = Document(docments_path)
            total_count = 0
            month_count = 0
            month_info = {}
            actual_records = 0
            actual_records_month = {}
            wrong_records_count = 0
            start_count = False
            for paragraph in doc.paragraphs:
                if start_count:
                    if paragraph.text == "":
                        break
                    if not paragraph.text[0].isdigit() and not paragraph.text.endswith("："):
                        total_count += 1
                        month_count += 1
                        month_info[month] = month_count
                    else:
                        month_count = 0
                        month = int(re.search(r"(\d+)月", paragraph.text).group(1))
                        actual_records_month[month] = int(re.search(r"共(\d+)个", paragraph.text).group(1)) if re.search(r"共(\d+)个", paragraph.text) else 0
                if not start_count and "新事物（或重大事件）" in paragraph.text:
                    start_count = True
                    actual_records = int(re.search(r"(\d+)个新事物（或重大事件）", paragraph.text).group(1)) if re.search(r"(\d+)个新事物（或重大事件）", paragraph.text) else 0
            print(f"在{year}年的年度总结中共发现{total_count}个新事物（或重大事件）")
            log(f"在{year}年的年度总结中共发现{total_count}个新事物（或重大事件）", "info", logfile_only=True)
            for k, v in month_info.items():
                if v == actual_records_month[k]:
                    print(f"{str(k)}月: {v}个")
                    log(f"{str(k)}月: {v}个", "info", logfile_only=True)
                else:
                    print(f"{str(k)}月: {v}个（年度总结中记录为{actual_records_month[k]}个，请更正）")
                    log(f"{str(k)}月: {v}个（年度总结中记录为{actual_records_month[k]}个，请更正）", "info", logfile_only=True)
                    wrong_records_count += 1
            if total_count == actual_records:
                print(f"在年度总结中共发现{total_count}个新事物（或重大事件），与实际记录一致")
                log(f"在年度总结中共发现{total_count}个新事物（或重大事件），与实际记录一致", "info", logfile_only=True)
            else:
                print(f"在年度总结中共发现{total_count}个新事物（或重大事件），但年度总结中记录为{actual_records}个，请更正")
                log(f"在年度总结中共发现{total_count}个新事物（或重大事件），但年度总结中记录为{actual_records}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            print(f"对{year}年新事物（或重大事件）的统计与检查已完成，共发现{wrong_records_count}项记录错误")
            log(f"对{year}年新事物（或重大事件）的统计与检查已完成，共发现{wrong_records_count}项记录错误", "info", logfile_only=True)
            print()
            log("", "info", logfile_only=True)
        elif args.split(" ")[0] == "normal_period":
            year = args.split(" ")[2]
            docments_path = os.path.join(ftfpath, year, "年度总结.docx")
            if not os.path.exists(docments_path):
                print(f"未找到{year}年的年度总结")
                log(f"未找到{year}年的年度总结", "warning", logfile_only=True)
                return
            doc = Document(docments_path)
            total_count = 0
            total_count_strong = 0
            total_count_weak = 0
            actual_total_count_strong = 0
            actual_total_count_weak = 0
            month_count = 0
            month_info = {}
            actual_records = 0
            wrong_records_count = 0
            start_count = False
            for paragraph in doc.paragraphs:
                if start_count:
                    if paragraph.text == "":
                        break
                    if not paragraph.text[0].isdigit() and not paragraph.text.endswith("："):
                        total_count += 1
                        month_count += 1
                        month_info[month] = month_count
                    else:
                        month_count = 0
                        month = int(re.search(r"(\d+)月", paragraph.text).group(1))
                    if "时期（强）" in paragraph.text:
                        total_count_strong += 1
                    if "时期（弱）" in paragraph.text:
                        total_count_weak += 1
                if not start_count and "常规“时期”" in paragraph.text:
                    start_count = True
                    actual_records = int(re.search(r"(\d+)个常规“时期”", paragraph.text).group(1)) if re.search(r"(\d+)个常规“时期”", paragraph.text) else 0
                    actual_total_count_strong = int(re.search(r"(\d+)个强“时期”", paragraph.text).group(1)) if re.search(r"(\d+)个强“时期”", paragraph.text) else 0
                    actual_total_count_weak = int(re.search(r"(\d+)个弱“时期”", paragraph.text).group(1)) if re.search(r"(\d+)个弱“时期”", paragraph.text) else 0
            print(f"在{year}年的年度总结中共发现{total_count}个常规“时期”")
            log(f"在{year}年的年度总结中共发现{total_count}个常规“时期”", "info", logfile_only=True)
            print(f"其中强“时期”有{total_count_strong}个，弱“时期”有{total_count_weak}个")
            log(f"其中强“时期”有{total_count_strong}个，弱“时期”有{total_count_weak}个", "info", logfile_only=True)
            for k, v in month_info.items():
                print(f"{str(k)}月: {v}个")
                log(f"{str(k)}月: {v}个", "info", logfile_only=True)
            if total_count == actual_records:
                print(f"在年度总结中共发现{total_count}个常规“时期”，与实际记录一致")
                log(f"在年度总结中共发现{total_count}个常规“时期”，与实际记录一致", "info", logfile_only=True)
            else:
                print(f"在年度总结中共发现{total_count}个常规“时期”，但年度总结中记录为{actual_records}个，请更正")
                log(f"在年度总结中共发现{total_count}个常规“时期”，但年度总结中记录为{actual_records}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            if total_count_strong == actual_total_count_strong:
                print(f"发现强“时期”{total_count_strong}个，与实际记录一致")
                log(f"发现强“时期”{total_count_strong}个，与实际记录一致", "info", logfile_only=True)
            else:
                print(f"发现强“时期”{total_count_strong}个，但年度总结中记录为{actual_total_count_strong}个，请更正")
                log(f"发现强“时期”{total_count_strong}个，但年度总结中记录为{actual_total_count_strong}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            if total_count_weak == actual_total_count_weak:
                print(f"发现弱“时期”{total_count_weak}个，与实际记录一致")
                log(f"发现弱“时期”{total_count_weak}个，与实际记录一致", "info", logfile_only=True)
            else:
                print(f"发现弱“时期”{total_count_weak}个，但年度总结中记录为{actual_total_count_weak}个，请更正")
                log(f"发现弱“时期”{total_count_weak}个，但年度总结中记录为{actual_total_count_weak}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            print(f"对{year}年常规“时期”的统计与检查已完成，共发现{wrong_records_count}项记录错误")
            log(f"对{year}年常规“时期”的统计与检查已完成，共发现{wrong_records_count}项记录错误", "info", logfile_only=True)
            print()
            log("", "info", logfile_only=True)
        elif args.split(" ")[0] == "combined_period":
            year = args.split(" ")[2]
            docments_path = os.path.join(ftfpath, year, "年度总结.docx")
            if not os.path.exists(docments_path):
                print(f"未找到{year}年的年度总结")
                log(f"未找到{year}年的年度总结", "warning", logfile_only=True)
                return
            doc = Document(docments_path)
            total_count = 0
            total_count_strong = 0
            total_count_weak = 0
            actual_total_count_strong = 0
            actual_total_count_weak = 0
            month_info = []
            actual_records = 0
            wrong_records_count = 0
            start_count = False
            for paragraph in doc.paragraphs:
                if start_count:
                    if paragraph.text == "":
                        break
                    total_count += 1
                    month_info.append(re.search(r"\d+~\d+月", paragraph.text).group(0))
                    if re.search(r"强.*“时期”", paragraph.text):
                        total_count_strong += 1
                    if re.search(r"弱.*“时期”", paragraph.text):
                        total_count_weak += 1
                if not start_count and "合称“时期”" in paragraph.text:
                    start_count = True
                    actual_records = int(re.search(r"(\d+)个合称“时期”", paragraph.text).group(1)) if re.search(r"(\d+)个合称“时期”", paragraph.text) else 0
                    actual_total_count_strong = int(re.search(r"(\d+)个强“时期”", paragraph.text).group(1)) if re.search(r"(\d+)个强“时期”", paragraph.text) else 0
                    actual_total_count_weak = int(re.search(r"(\d+)个弱“时期”", paragraph.text).group(1)) if re.search(r"(\d+)个弱“时期”", paragraph.text) else 0
            if total_count == actual_records:
                print(f"在{year}年的年度总结中共发现{total_count}个合称“时期”: {', '.join([i for i in month_info])}")
                log(f"在{year}年的年度总结中共发现{total_count}个合称“时期”: {', '.join([i for i in month_info])}", "info", logfile_only=True)
            else:
                print(f"在{year}年的年度总结中共发现{total_count}个合称“时期”，但年度总结中记录为{actual_records}个，请更正")
                log(f"在{year}年的年度总结中共发现{total_count}个合称“时期”，但年度总结中记录为{actual_records}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            if total_count_strong == actual_total_count_strong:
                print(f"发现强“时期”{total_count_strong}个，与实际记录一致")
                log(f"发现强“时期”{total_count_strong}个，与实际记录一致", "info", logfile_only=True)
            else:
                print(f"发现强“时期”{total_count_strong}个，但年度总结中记录为{actual_total_count_strong}个，请更正")
                log(f"发现强“时期”{total_count_strong}个，但年度总结中记录为{actual_total_count_strong}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            if total_count_weak == actual_total_count_weak:
                print(f"发现弱“时期”{total_count_weak}个，与实际记录一致")
                log(f"发现弱“时期”{total_count_weak}个，与实际记录一致", "info", logfile_only=True)
            else:
                print(f"发现弱“时期”{total_count_weak}个，但年度总结中记录为{actual_total_count_weak}个，请更正")
                log(f"发现弱“时期”{total_count_weak}个，但年度总结中记录为{actual_total_count_weak}个，请更正", "info", logfile_only=True)
                wrong_records_count += 1
            print(f"对{year}年合称“时期”的统计与检查已完成，共发现{wrong_records_count}项记录错误")
            log(f"对{year}年合称“时期”的统计与检查已完成，共发现{wrong_records_count}项记录错误", "info", logfile_only=True)
            print()
            log("", "info", logfile_only=True)
        else:
            print(self.do_count.__doc__)

    def complete_count(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        if re.match(r"count (N&M|normal_period|combined_period) in \w*", line):
            return [i for i in self.years if i.startswith(text)]
        if len(line.split(" ")) > 2:
            return ["in"]
        if line.startswith("count "):
            return [i for i in ["N&M", "normal_period", "combined_period"] if i.startswith(text)]
        return []
    
    def do_calculate(self, args: str):
        """
        计算“时期”的强度。

        语法：calculate [strong <count>] [weak <count>] [scattered <coefficient>] start <time> end <time>/now [/?]
            strong <count>              强定位物的数量。
            weak <count>                弱定位物的数量。
            scattered <coefficient>     若为分散性“时期”，指定的分散系数，否则默认为1。
            start <month>               “时期”记录起算时间，格式为“年份/月份”，例如“2025/1”。
            end <month>/now             指定结束时间，或键入“now”以选择当前时间。
            /?                          显示此帮助文档。
        
        注：以上输入参数的顺序可以任意调整，但起算时间必须早于结束时间。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_calculate.__doc__)
            return
        parts = args.split(" ")
        if "start" not in parts or "end" not in parts:
            print(self.do_calculate.__doc__)
            return
        start_index = parts.index("start")
        end_index = parts.index("end")
        if start_index > end_index:
            print(self.do_calculate.__doc__)
            return
        start_time = parts[start_index + 1]
        end_time = parts[end_index + 1]
        strong_count = int(parts[parts.index("strong") + 1]) if "strong" in parts else 0
        weak_count = int(parts[parts.index("weak") + 1]) if "weak" in parts else 0
        scattered_coefficient = int(parts[parts.index("scattered") + 1]) if "scattered" in parts else 1
        start_year, start_month = map(int, start_time.split("/"))
        if end_time == "now":
            now = datetime.datetime.now()
            end_year, end_month = now.year, now.month
            end_time = f"{end_year}/{end_month}"
        else:
            end_year, end_month = map(int, end_time.split("/"))
        if (start_year > end_year) or (start_year == end_year and start_month > end_month):
            print("结束时间必须晚于起算时间")
            log("结束时间必须晚于起算时间", "warning", logfile_only=True)
            return
        interval_months = (end_year - start_year) * 12 + (end_month - start_month)
        total_items = strong_count + weak_count
        total_strength = strong_count * 2 + weak_count * 1
        intensity = ((total_items * total_strength) / scattered_coefficient) * math.exp(-0.1 * interval_months)
        print(f"从{start_time}到{end_time}，间隔{interval_months}个月")
        print(f"强定位物数量: {strong_count}")
        print(f"弱定位物数量: {weak_count}")
        print(f"分散系数: {scattered_coefficient}")
        print(f"计算得到的“时期”强度为: {intensity:.2f}夕")
        log(f"计算得到的“时期”强度为: {intensity:.2f}夕", "info", logfile_only=True)

    def complete_calculate(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        options = ["strong", "weak", "scattered", "start", "end"]
        parts = line.split(" ")
        if len(parts) > 1:
            if parts[-1] in options:
                return []
            elif parts[-2] == "start":
                if "/" in text:
                    parts = text.split("/")
                    if len(parts) == 2:
                        year_part, month_part = parts
                        if year_part in self.years or any(y.startswith(year_part) for y in self.years):
                            months = [str(i) for i in range(1, 13)]
                            completions = [f"{year_part}/{m}" for m in months if m.startswith(month_part)]
                            return completions
                    return []
                else:
                    year_completions = [y for y in self.years if y.startswith(text)]
                    return [f"{y}/" for y in year_completions]
            elif parts[-2] == "end":
                if "/" in text:
                    parts = text.split("/")
                    if len(parts) == 2:
                        year_part, month_part = parts
                        if year_part in self.years or any(y.startswith(year_part) for y in self.years):
                            months = [str(i) for i in range(1, 13)]
                            completions = [f"{year_part}/{m}" for m in months if m.startswith(month_part)]
                            return completions
                    return []
                else:
                    if text.startswith("n"):
                        return ["now"]
                    year_completions = [f"{y}/" for y in self.years if y.startswith(text)] + ["now"]
                    return [y for y in year_completions if y.startswith(text)]
            else:
                return [i for i in options if i.startswith(parts[-1])]
        else:
            return options

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
        c = choice("YN", Fore.LIGHTYELLOW_EX + "协议创始人，请牢记您的命令一旦执行便无法撤回，您确定要执行此命令吗")
        os.system("color 0c")
        if c == 2:
            return
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
        os.system("color 0a")
        raise CommandLineExit()
    
    def do_dellog(self, args: str):
        """
        删除日志文件。

        语法：dellog <logname/*> [/?]
            logname     指定删除的日志文件的名称。
            *           表示所有日志文件。
            /?          显示此帮助文档。
        """
        if args.split(" ")[0] == "/?" or args == "":
            print(self.do_dellog.__doc__)
            return
        logname = args
        if logname == "*":
            for i in glob("logs\\*.log"):
                os.remove(i)
                print(f"已删除{i}")
                log(f"已删除{i}", "info", logfile_only=True)
        else:
            if not os.path.exists(f"logs\\{logname}.log"):
                print(f"未找到{logname}的日志文件")
                log(f"未找到{logname}的日志文件", "warning", logfile_only=True)
                return
            os.remove(f"logs\\{logname}.log")
            print(f"已删除{logname}.log")
            log(f"已删除{logname}.log", "info", logfile_only=True)

    def complete_dellog(self, text: str, line: str, begidx: int, endidx: int) -> list[str]:
        if re.match(r"dellog \w*", line):
            return [i.replace("logs\\", "").replace(".log", "") for i in glob("logs\\*.log") if i.replace("logs\\", "").replace(".log", "").startswith(text)]

def help_ftf_admin():
    print(Fore.LIGHTRED_EX + "您已处在协议创始人权限下")
