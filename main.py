from pywinauto.application import Application
from libs import *
import sys
import cv2

os.system("title FTF v2.10.1")

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
            log("未配置《朝花夕拾协议》根目录，请前往config.ini中配置ftfpath", "warning")
            raise UserWarning(
                "未配置《朝花夕拾协议》根目录，请前往config.ini中配置ftfpath"
            )
        elif not os.path.exists(config.ftfpath):
            log("《朝花夕拾协议》根目录无效，请确认路径是否正确，然后前往config.ini中修改ftfpath", "warning")
            raise UserWarning(
                "《朝花夕拾协议》根目录无效，请确认路径是否正确，然后前往config.ini中修改ftfpath"
            )

        result = choice("123456789q", "请选择你要使用的功能:", hide=True)
        if result == 1:
            log("正在打开《朝花夕拾协议》", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/朝花夕拾协议.docx")
        elif result == 2:
            log("正在打开《朝花夕拾协议》试题", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/《朝花夕拾协议》熟悉程度统一考试.docx")
        elif result == 3:
            log("正在打开《朝花夕拾协议》试题答案", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/《朝花夕拾协议》熟悉程度统一考试（答案）.docx")
        elif result == 4:
            log("正在打开《朝花夕拾协议通用掩盖指南》", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/朝花夕拾协议通用掩盖指南.docx")
        elif result == 5:
            log("正在打开《朝花夕拾协议消极情绪应对指南》", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/朝花夕拾协议消极情绪应对指南.docx")
        elif result == 6:
            log("正在打开《朝花夕拾协议》主题曲《记忆的树林》", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/《朝花夕拾协议》主题曲《记忆的树林》.docx")
        elif result == 7:
            log("正在打开《朝花夕拾协议》印象曲《寻忆》", "info", logfile_only=True)
            os.startfile(f"{config.ftfpath}/《朝花夕拾协议》印象曲《寻忆》.docx")
        elif result == 8:
            try:
                log("《朝花夕拾协议》命令行启动", "info", logfile_only=True)
                FTF_cmd.cmdloop()
            except CommandLineExit:
                continue
            except AdminMode:
                try:
                    FTF_ADMIN_cmd.cmdloop()
                except CommandLineExit:
                    continue
        elif result == 9:
            break
        else:
            sys.exit(0)

    os.system("cls")

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
    log(f"连接到控制器: {config.controller}", "info")
    log("开始监听", "info")
    say_in_english("start listening")

    try:
        executant_window = wechat_window.child_window(
            title=config.controller, control_type="ListItem")
        executant_wrapper_object = executant_window.wrapper_object()
        command_list = ["/test", "/shutdown", "/open-url", "/exec", "/send-file", "/screenshot", "/camera-screenshot", "/exit", "/transfer"]
        while True:
            for i in executant_wrapper_object.descendants():
                if i.window_text().split(" ")[0] in command_list:
                    log(f"ControlType: {i.friendly_class_name()}", "debug")
                    command: str = i.window_text()
                    log(f"[远程终端指令] {command}", "info")
                    say_in_english(f"command received: {command.split(" ")[0][1:]}")
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
                        log(f"获取到路径/网址: {url}", "info")
                        say_in_english("starting url")
                        log("正在打开路径/网址", "info")
                        wechat(f"已打开路径/网址{url}", executant_wrapper_object)
                        wechat_window.minimize()
                        os.system(f"start {url}")
                    elif command.split(" ")[0] == "/exec":
                        say_in_english("remote command execution detected")
                        remote_cmd = ""
                        for i in command.split(" ")[1:]:
                            remote_cmd += i + " "
                        log(f"检测到远程执行命令: {remote_cmd}", "info")
                        say_in_english("executing command")
                        log("开始执行", "info")
                        try:
                            exec(remote_cmd)
                        except Exception as e:
                            log("执行中发生错误", "warning")
                            log(str(e), "error")
                            wechat(str(e), executant_wrapper_object)
                            wechat_window.minimize()
                        wechat("执行完毕", executant_wrapper_object)
                        wechat_window.minimize()
                    elif command.split(" ")[0] == "/send-file":
                        say_in_english("file sending command detected")
                        log("检测到文件发送指令", "info")
                        say_in_english("getting file")
                        log("开始获取文件", "info")
                        copyfile(*command.split(" ")[1:])
                        log("开始发送", "info")
                        wechat("开始发送", executant_wrapper_object)
                        pyautogui.hotkey("ctrl", "v")
                        pyautogui.hotkey("enter")
                        wechat_window.minimize()
                    elif command == "/screenshot":
                        say_in_english("screenshot command detected")
                        log("检测到截屏指令", "info")
                        pyautogui.screenshot().save("screenshot.jpg")
                        say_in_english("sending screenshot")
                        log("发送截屏中", "info")
                        copyfile("screenshot.jpg")
                        wechat("发送截屏中", executant_wrapper_object)
                        pyautogui.hotkey("ctrl", "v")
                        pyautogui.hotkey("enter")
                        wechat_window.minimize()
                        os.remove("screenshot.jpg")
                    elif command == "/camera-screenshot":
                        say_in_english("camera screenshot command detected")
                        log("检测到摄像头截屏指令", "info")
                        cap = cv2.VideoCapture(0)
                        if not cap.isOpened():
                            say_in_english("unable to open camera")
                            log("无法打开摄像头", "warning")
                            wechat("无法打开摄像头", executant_wrapper_object)
                            wechat_window.minimize()
                            continue
                        ret, frame = cap.read()
                        if not ret:
                            say_in_english("unable to capture image from camera")
                            log("无法从摄像头捕获图像", "warning")
                            wechat("无法从摄像头捕获图像", executant_wrapper_object)
                            wechat_window.minimize()
                            cap.release()
                            continue
                        cv2.imwrite("camera_screenshot.jpg", frame)
                        cap.release()
                        say_in_english("sending camera screenshot")
                        log("发送摄像头截屏中", "info")
                        copyfile("camera_screenshot.jpg")
                        wechat("发送摄像头截屏中", executant_wrapper_object)
                        pyautogui.hotkey("ctrl", "v")
                        pyautogui.hotkey("enter")
                        wechat_window.minimize()
                        os.remove("camera_screenshot.jpg")
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
                            "远程终端要求将控制权限转交本地终端。若要重新将权限移交远程终端，请在监听终端命令行中使用restart指令重启终端")
                else:
                    if i.window_text().split(" ")[0].startswith("/"):
                        wechat("没有此命令", executant_wrapper_object)
                        wechat_window.minimize()
            time.sleep(1)
    except Exception as e:
        if type(e) is TransferTerminalControl:
            log(str(e), "info")
        else:
            log(str(e), "error")
            log("监听中发生错误，详细信息已被写入日志", "exception")
            say_in_english("an error was caught")
            print(Fore.BLUE)
        is_continue = False
        while True:
            result = input("[FTF Terminal Listener] ")
            if result.isspace() or result == "":
                break
            log(f"[监听终端指令] {result}", "info")
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
