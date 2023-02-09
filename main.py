import pyautogui
from pywinauto.application import Application
from libs import *

log("终端启动中", "info")
log("正在连接到微信", "info")
say_in_english("starting terminal")
say_in_english("connecting")

pyautogui.hotkey("ctrl", "alt", "w")
app = Application(backend="uia").connect(process=get_wechat_pid())
wechat_window = app.window(class_name="WeChatMainWndForPC")
wechat_window.minimize()
pyautogui.FAILSAFE = False

log("已尝试连接，如没有成功，请手动启动微信", "info")
say_in_english("attempted to connect")
say_in_english("please stand by")

try:
    log("开始监听", "info")
    say_in_english("start listening")
    executant_window = wechat_window.child_window(title="一只叫迷迭香的菲林", control_type="ListItem")
    executant_wrapper_object = executant_window.wrapper_object()
    command_list = ["/start protocol", "/stop protocol", "/shutdown"]
    while True:
        for i in executant_wrapper_object.descendants():
            if i.friendly_class_name() == "Static":
                if i.window_text() in command_list:
                    command = i.window_text()
                    log("获取到指令: %s"%command, "info")
                    say_in_english("command received: %s"%command.replace("/", ""))
                    if command == "/start protocol":
                        say_in_english("protocol activation command detected")
                        log("检测到协议激活指令", "info")
                        log("协议已激活", "info")
                        pyautogui.hotkey("ctrl", "alt", "w")
                        executant_wrapper_object.click_input()
                        executant_wrapper_object.type_keys("[FTF] 协议已激活", with_spaces=True)
                        pyautogui.hotkey("enter")
                        wechat_window.minimize()
                    elif command == "/stop protocol":
                        say_in_english("protocol terminated command detected")
                        log("检测到协议终止指令", "info")
                        log("协议已终止", "info")
                        pyautogui.hotkey("ctrl", "alt", "w")
                        executant_wrapper_object.click_input()
                        executant_wrapper_object.type_keys("[FTF] 协议已终止", with_spaces=True)
                        pyautogui.hotkey("enter")
                        wechat_window.minimize()
                    elif command == "/shutdown":
                        say_in_english("shutdown command detected")
                        log("检测到关机指令", "info")
                        say_in_english("the computer will shutdown in T-minus ten seconds")
                        log("将在10秒后关机", "info")
                        pyautogui.hotkey("ctrl", "alt", "w")
                        executant_wrapper_object.click_input()
                        executant_wrapper_object.type_keys("[FTF] 将在10秒后关机", with_spaces=True)
                        pyautogui.hotkey("enter")
                        wechat_window.minimize()
                        os.system("shutdown -s -t 10")
        time.sleep(1)
except Exception as e:
    log("监听中发生错误，我们获取了以下信息", "warning")
    say_in_english("an error was caught")
    log(str(e), "error")
    input()
