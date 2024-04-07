import sys
from pywinauto.application import Application
from libs import *

os.system("title FTF v2.1.2")

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
        print("3. 启动《朝花夕拾协议》命令行")
        print("4. 启动《朝花夕拾协议》监听终端")
        result = choice("1234q", "请选择你要使用的功能:", hide=True)
        if result == 1:
            log("正在打开《朝花夕拾协议》", "info")
            os.system(
                "start https://docs.qq.com/doc/p/c9ef326c964ac8a2fea816ad59822ad3cba514f8")
        elif result == 2:
            log("正在打开《朝花夕拾协议》试题", "info")
            os.system(
                "start https://docs.qq.com/doc/p/d46f28744498dcd14313090c195ee1b629971f9f?u=ee618ac0d45149c5a407d1dcf3e9d78d")
        elif result == 3:
            try:
                log("《朝花夕拾协议》命令行启动", "info")
                FTF_cmd.cmdloop()
            except CommandLineExit:
                continue
            except AdminMode:
                try:
                    FTF_ADMIN_cmd.cmdloop()
                except CommandLineExit:
                    continue
        elif result == 4:
            break
        elif result == 5:
            sys.exit(0)

    log("监听终端启动中", "info")
    log("正在连接到微信", "info")
    say_in_english("starting terminal")
    say_in_english("connecting")

    pyautogui.hotkey("ctrl", "alt", "w")
    app = Application(backend="uia").connect(process=get_wechat_pid())
    wechat_window = app.window(class_name="WeChatMainWndForPC")
    wechat_window.minimize()
    pyautogui.FAILSAFE = False

    log("已尝试连接，若没有成功，请手动启动微信", "info")
    say_in_english("attempted to connect")
    say_in_english("please stand by")
    log("开始监听", "info")
    say_in_english("start listening")

    try:
        executant_window = wechat_window.child_window(
            title="一只叫迷迭香的菲林", control_type="ListItem")
        executant_wrapper_object = executant_window.wrapper_object()
        command_list = ["/test", "/shutdown", "/open-url", "/exit", "/transfer"]
        while True:
            for i in executant_wrapper_object.descendants():
                if i.window_text().split(" ")[0] in command_list:
                    log("ControlType: %s" % i.friendly_class_name(), "debug")
                    command: str = i.window_text()
                    log("[远程终端指令] %s" % command, "info")
                    say_in_english("command received: %s" % command[1:])
                    if command == "/test":
                        say_in_english("protocol activation command detected")
                        log("检测到测试指令", "info")
                        log("一切为了不远后的旧事重提", "info")
                        wechat("一切为了不远后的旧事重提", executant_wrapper_object)
                        wechat_window.minimize()
                    elif command == "/shutdown":
                        say_in_english("shutdown command detected")
                        log("检测到关机指令", "info")
                        say_in_english(
                            "the computer will shutdown in T-minus ten seconds")
                        log("将在10秒后关机", "info")
                        wechat("将在10秒后关机", executant_wrapper_object)
                        wechat_window.minimize()
                        os.system("shutdown -s -t 10")
                    elif command.startswith("/open-url"):
                        say_in_english("url open command detected")
                        log("检测到路径/网址启动指令", "info")
                        url = ""
                        for i in command.split(" ")[1:]:
                            url += i + " "
                        say_in_english("getting url")
                        log("获取到路径/网址: %s" % url, "info")
                        say_in_english("starting url")
                        log("正在打开路径/网址", "info")
                        wechat("已打开路径/网址%s" % url, executant_wrapper_object)
                        wechat_window.minimize()
                        os.system("start %s" % url)
                    elif command == "/exit":
                        say_in_english("exit command detected")
                        log("检测到终端退出指令", "info")
                        say_in_english("terminal listening task terminated")
                        log("终端监听任务终止", "info")
                        wechat("终端已关闭", executant_wrapper_object)
                        wechat_window.minimize()
                        exit(0)
                    elif command == "/transfer":
                        say_in_english("transfer command detected")
                        log("检测到终端控制方式更改指令", "info")
                        say_in_english("transfering terminal control")
                        log("正在更改终端控制方式", "info")
                        wechat(
                            "终端控制方式已由远程终端控制更改为本地终端控制，你所有的远程终端操作权限已被转移至本地终端", executant_wrapper_object)
                        wechat_window.minimize()
                        raise TransferTerminalControl(
                            "远程终端要求将控制权限转为本地终端。若要重新将权限移交远程终端，请在本地终端中使用restart指令重启终端")
            time.sleep(1)
    except Exception as e:
        if type(e) is TransferTerminalControl:
            log(str(e), "info")
        else:
            log("监听中发生错误，我们获取了以下信息", "warning")
            say_in_english("an error was caught")
            log(str(e), "error")
            print(Fore.BLUE)
        is_continue = False
        while True:
            result = input("[FTF Terminal Listener] ")
            if result.isspace() or result == "":
                break
            log("[监听终端指令] %s" % result, "info")
            if result == "restart":
                log("重启终端", "info")
                say_in_english("restarting terminal")
                is_continue = True
                break
            log("未知指令", "warning")
            continue
        if is_continue:
            continue
        break
