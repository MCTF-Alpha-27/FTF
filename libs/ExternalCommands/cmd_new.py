from libs.command_lines import *

cmd_template = """from libs.command_lines import *

def {0}(self, cmd: str):
{1}
FTFCmd.do_{0} = {0}
"""

def new(self, cmd: str):
    """创建一个新的外部命令"""
    cmd = cmd.split(" ")[0]
    if cmd == "/?":
        print(new.__doc__)
        return
    if os.path.exists("libs\\ExternalCommands\\cmd_%s.py"%cmd):
        log("该外部命令已经存在，请更改您的外部命令名称", "warning")
        return
    with open("libs\\ExternalCommands\\cmd_%s.py"%cmd, "w", encoding="utf-8") as f:
        func = ""
        new_cmd = ""
        log("创建外部命令: libs\\ExternalCommands\\cmd_%s.py"%cmd, "info")
        log("请为新命令的函数编程，输入#end来结束编程", "info")
        while func != "#end":
            func = input("... ")
            log("... " + func, "info")
            new_cmd += "    " + func + "\n"
        new_cmd = cmd_template.format(cmd, new_cmd)
        f.write(new_cmd)
    log("新命令已保存，重启命令行后生效", "info")

FTFCmd.do_new = new
