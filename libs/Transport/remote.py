"""
与具体传输通道无关的监听终端命令处理
"""

import os

import cv2
import pyautogui

from libs.AI.agent import Agent, AgentReply
from libs.CLI import api
from libs.Common import config
from libs.Common.functions import log

# 命令行命令的远程入口
CLI_COMMANDS = {
    "/find": "find",
    "/count": "count",
    "/periods": "periods",
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

# AI 上下文：每个会话保留最近几轮问答，用于理解“刚才”“2026年”这类追问
MAX_HISTORY_MESSAGES = 12
MAX_HISTORY_CHARS = 800

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
    # AI 语义解析器，首次使用时初始化
    _agent: Agent | None = None
    _agent_ready = False
    # 各会话的 AI 上下文：会话标识 -> [{"role": ..., "content": ...}]
    _histories: dict[str, list[dict[str, str]]] = {}

    def conversation_key(self) -> str:
        """当前会话标识；同一标识的连续消息共享 AI 上下文"""
        return ""

    def _history_snapshot(self) -> list[dict[str, str]]:
        """取出当前会话的历史副本，交给 AI 使用"""
        return list(self._histories.get(self.conversation_key(), []))

    def _remember_turn(self, request: str, answer: str) -> None:
        """记录一轮问答，并按消息条数裁剪旧上下文"""
        request = (request or "").strip()
        answer = (answer or "").strip()
        if not request or not answer:
            return
        history = self._histories.setdefault(self.conversation_key(), [])
        history.append(
            {"role": "user", "content": request[:MAX_HISTORY_CHARS]}
        )
        history.append(
            {"role": "assistant", "content": answer[:MAX_HISTORY_CHARS]}
        )
        limit = MAX_HISTORY_MESSAGES - MAX_HISTORY_MESSAGES % 2
        if len(history) > limit:
            del history[: len(history) - limit]

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
        shown = (
            f"{first} ***"
            if lowered == "/auth" and " " in command
            else command
        )
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
            # 普通文本交给 AI 语义解析；未启用 AI 时保持静默
            return self._run_ai(command)

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

    def _ai_agent(self) -> Agent | None:
        """按需创建 AI 语义解析器；未启用或未配置密钥时返回 None"""
        if self._agent_ready:
            return self._agent
        self._agent_ready = True
        if not config.ai_enabled:
            log(
                "AI 语义解析未启用（config.ini 的 [AI] enabled=False）", "info"
            )
            return None
        try:
            agent = Agent()
        except Exception as e:
            log(f"AI 语义解析初始化失败: {e}", "error")
            return None
        if not agent.available:
            log("AI 语义解析不可用：未配置硅基流动 API 密钥", "warning")
            return None
        self._agent = agent
        log(f"AI 语义解析已启用: {agent.client.model}", "info")
        return agent

    def _run_ai(self, text: str) -> str:
        """用 AI 把自然语言请求转换成命令行命令

        创始人工具仍会走密钥卡片与 /y 确认，AI 只负责选出命令
        """
        agent = self._ai_agent()
        if agent is None:
            return "continue"

        history = self._history_snapshot()
        if history:
            log(f"[AI 上下文] 注入 {len(history)} 条历史消息", "info")

        try:
            reply = agent.handle(text, history=history)
        except Exception as e:
            log(f"[AI 语义解析失败] {e}", "error")
            self._remember_turn(text, f"（AI 解析失败: {e}）")
            self.notify("AI 解析失败，请改用 /find 等手动命令")
            return "continue"

        self._remember_turn(text, self._answer_text(reply))

        if reply.needs_permission:
            log(
                f"[AI 计划需要创始人权限] {reply.command}",
                "info",
                logfile_only=True,
            )
            self.notify(
                f"识别为创始人操作: {reply.command}\n请先输入密钥，再重新发送这条请求"
            )
            self._replace_auth_card()
            return "continue"

        if reply.needs_confirmation:
            # 创始人操作尚未执行；公开命令的内部确认已经执行过一遍
            return self._ask(
                reply.command,
                reply.question,
                auto_answer=reply.result is not None,
            )

        if not reply.command:
            self.notify(reply.text or reply.error or "没有理解这条请求")
            return "continue"

        if not reply.ok:
            self.notify(
                f"识别命令: {self._command_text(reply)}\n"
                f"{reply.error or '执行失败'}"
            )
            return "continue"

        if reply.file_path:
            self._send_ai_file(reply)
            self._refresh_auth_card()
            return "continue"

        log(f"[AI 执行命令] {self._command_text(reply)}", "info")
        self._send_ai_result(reply)
        self._refresh_auth_card()
        return "continue"

    @staticmethod
    def _command_text(reply: AgentReply) -> str:
        """把一次或多次工具调用整理成用于展示的命令行"""
        commands = [c for c in reply.commands if c] or [reply.command]
        return "\n".join(command for command in commands if command)

    @staticmethod
    def _answer_text(reply: AgentReply) -> str:
        """整理本轮回复，作为下一轮追问的历史上下文"""
        answer = (reply.summary or reply.text or reply.error or "").strip()
        command = RemoteTerminal._command_text(reply)
        if reply.needs_permission:
            note = f"（需要创始人密钥，未执行: {command}）"
        elif reply.needs_confirmation:
            note = f"（等待用户确认: {command}）"
        elif reply.file_path and reply.ok:
            note = f"（发送文件: {os.path.basename(reply.file_path)}）"
        elif command and reply.ok:
            note = f"（已执行命令: {command}）"
        else:
            note = ""
        if note:
            answer = f"{answer}\n{note}" if answer else note
        return answer or "（没有返回内容）"

    def _send_ai_result(self, reply: AgentReply) -> None:
        """回执 AI 的结果：单行直接发，多行把原始输出作为文件补充发送"""
        summary = (reply.summary or "").strip()
        output = (reply.text or "").strip()
        command = self._command_text(reply)
        if not output:
            self.notify(summary or "执行完毕")
            self.notify(f"命令: {command}")
            return
        if "\n" not in output:
            parts = [summary, output, f"命令: {command}"]
            self.notify("\n".join(part for part in parts if part))
            return
        if summary:
            self.notify(f"{summary}\n命令: {command}")
            # 摘要不比原始输出短时认为已覆盖全部条目，不再重复发送文件
            if len(summary.splitlines()) >= len(output.splitlines()):
                return
        self._send_result_file(f"命令: {command}\n\n{output}")

    def _send_ai_file(self, reply: AgentReply) -> None:
        """发送 AI 选中的文档；失败时告知用户"""
        path = reply.file_path
        name = os.path.basename(path)
        log(f"[AI 发送文件] {name}", "info")
        try:
            self.send_file(path)
        except Exception as e:
            log(f"[AI 发送文件失败] {e}", "error")
            self.notify(f"发送文件失败: {e}")
            return
        self.notify(f"已发送文件: {name}")

    def _refresh_auth_card(self) -> None:
        """未鉴权时把密钥卡片挪到最新消息下方"""
        if api.is_admin() or not self._auth_card_id:
            return
        self._replace_auth_card()

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

    def _run_cli(
        self, command_line: str, *, answer: bool | None = None
    ) -> str:
        """执行 api.py 封装好的命令行命令，并按输出长度选择回执方式"""
        log(f"执行命令行命令: {command_line}", "info")
        result = api.run(command_line, answer=answer)

        # 命令内部还有确认步骤时，交回用户决定
        if result.needs_confirmation:
            return self._ask(result.command, result.question, auto_answer=True)

        if not result.ok:
            log(f"命令行命令执行失败: {result.error}", "error")
            self._remember_turn(
                command_line, f"（执行失败: {result.error or '未知错误'}）"
            )
            self.notify(result.error or "执行失败")
            return "continue"

        output = result.output.strip()
        self._remember_turn(command_line, output or "（执行完毕，没有输出）")
        if not output:
            self.notify("执行完毕")
        elif "\n" in output:
            self._send_result_file(output)
        else:
            self.notify(output)
        return "continue"

    def _ask(
        self, command_line: str, question: str, *, auto_answer: bool
    ) -> str:
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
        return self._run_cli(
            command_line, answer=True if auto_answer else None
        )

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
