"""
《朝花夕拾协议》主程序
"""
import os
import sys

from colorama import Fore
from libs import config, listener
from libs.command_lines import FTFAdminCmd, FTFCmd, help_ftf, help_ftf_admin
from libs.exceptions import AdminMode, CommandLineExit, ConfigError
from libs.functions import choice, log

os.system("title FTF v3.0.0")

FTF_cmd = FTFCmd()
FTF_cmd.help_ftf = help_ftf
FTF_ADMIN_cmd = FTFAdminCmd()
FTF_ADMIN_cmd.help_ftf = help_ftf_admin

while True:
    while True:
        os.system("cls")
        
        print("欢迎使用《朝花夕拾协议》终端\n按下q键以退出\n")
        print("1. 《朝花夕拾协议》是什么")
        print("2. 《朝花夕拾协议》试题")
        print("3. 《朝花夕拾协议》试题答案")
        print("4. 《朝花夕拾协议通用掩盖指南》")
        print("5. 《朝花夕拾协议消极情绪应对指南》")
        print("6. 《朝花夕拾协议》主题曲《记忆的树林》")
        print("7. 《朝花夕拾协议》印象曲《寻忆》")
        print("8. 启动《朝花夕拾协议》命令行")
        print("9. 启动《朝花夕拾协议》监听终端")

        if config.ftfpath == r"{ftfpath}":
            log("未配置《朝花夕拾协议》根目录，请前往config.ini中配置ftfpath", "error")
            raise ConfigError(
                "未配置《朝花夕拾协议》根目录，请前往config.ini中配置ftfpath"
            )
        elif not os.path.exists(config.ftfpath):
            log("《朝花夕拾协议》根目录无效，请确认路径是否正确，然后前往config.ini中修改ftfpath", "error")
            raise ConfigError(
                "《朝花夕拾协议》根目录无效，请确认路径是否正确，然后前往config.ini中修改ftfpath"
            )

        result = choice("123456789q", "请选择你要使用的功能:", hide=True)
        if result == 1:
            log("正在打开《朝花夕拾协议》", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "朝花夕拾协议.docx"))
        elif result == 2:
            log("正在打开《朝花夕拾协议》试题", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "《朝花夕拾协议》熟悉程度统一考试.docx"))
        elif result == 3:
            log("正在打开《朝花夕拾协议》试题答案", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "《朝花夕拾协议》熟悉程度统一考试（答案）.docx"))
        elif result == 4:
            log("正在打开《朝花夕拾协议通用掩盖指南》", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "朝花夕拾协议通用掩盖指南.docx"))
        elif result == 5:
            log("正在打开《朝花夕拾协议消极情绪应对指南》", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "朝花夕拾协议消极情绪应对指南.docx"))
        elif result == 6:
            log("正在打开《朝花夕拾协议》主题曲《记忆的树林》", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "《朝花夕拾协议》主题曲《记忆的树林》.docx"))
        elif result == 7:
            log("正在打开《朝花夕拾协议》印象曲《寻忆》", "info", logfile_only=True)
            os.startfile(os.path.join(config.ftfpath, "《朝花夕拾协议》印象曲《寻忆》.docx"))
        elif result == 8:
            try:
                log("《朝花夕拾协议》命令行启动", "info", logfile_only=True)
                FTF_cmd.cmdloop()
            except CommandLineExit:
                continue
            except AdminMode:
                try:
                    FTF_ADMIN_cmd.intro = Fore.LIGHTRED_EX + f"欢迎您，协议创始人\n{FTF_cmd.cipher}\n"
                    FTF_ADMIN_cmd.cmdloop()
                except CommandLineExit:
                    continue
        elif result == 9:
            log("《朝花夕拾协议》监听终端启动", "info", logfile_only=True)
            listener.main()
        else:
            sys.exit(0)
