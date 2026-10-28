"""
《朝花夕拾协议》命令行 API

把 libs/CLI/ftf_cmd.py 中的命令行命令封装为可直接调用的函数，
供远程终端、AI 集成或其它脚本使用

用法示例：
    from libs.CLI.api import TerminalAPI

    terminal = TerminalAPI()                 # 默认只有公开命令权限
    result = terminal.run("find 愿望 in year 2025")

    if terminal.authenticate("……"):           # 用《朝花夕拾协议》的宗旨解锁创始人权限
        result = terminal.verify("2026")

    # 需要确认的命令会返回 needs_confirmation，由调用方决定后重试
    result = terminal.upload("2026/10")
    if result.needs_confirmation:
        result = terminal.upload("2026/10", answer=True)  # False 表示取消
"""

from __future__ import annotations

import contextlib
import inspect
import io
import re
import threading
from dataclasses import dataclass
from typing import Any, Iterator

from libs.CLI import ftf_cmd
from libs.CLI.ftf_cmd import FTFAdminCmd, FTFCmd, check_founder_key
from libs.Common.exceptions import (
    AdminMode,
    CommandLineExit,
    ConfirmationRequired,
)

_ANSI_PATTERN = re.compile(r"\x1b\[[0-9;]*m")

# 这些命令不提供 API 封装：open 依赖本机 GUI，exit 仅用于关闭命令行会话
_NOT_API = {"open", "exit"}


@dataclass(slots=True)
class CommandResult:
    """一次命令调用的结果"""

    command: str
    ok: bool
    output: str = ""
    error: str = ""
    exit_requested: bool = False
    needs_confirmation: bool = False
    question: str = ""

    def to_dict(self) -> dict[str, Any]:
        """转换为可 JSON 序列化的字典，便于直接回传给 AI 或前端"""
        return {
            "command": self.command,
            "ok": self.ok,
            "output": self.output,
            "error": self.error,
            "exit_requested": self.exit_requested,
            "needs_confirmation": self.needs_confirmation,
            "question": self.question,
        }


@dataclass(slots=True)
class CommandInfo:
    """命令的描述信息，用于构建 AI 工具清单"""

    name: str
    summary: str
    syntax: str
    doc: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "summary": self.summary,
            "syntax": self.syntax,
            "doc": self.doc,
        }


class TerminalAPI:
    """命令行命令的 API 封装

    默认只有公开命令（find / count / check / calculate）的权限
    调用 authenticate() 通过密钥校验后才会切换到创始人模式

    需要人工确认的命令不会自动执行：call 返回 needs_confirmation=True
    与 question，调用方确认后传 answer=True 重试，取消则传 answer=False
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._terminal: FTFCmd | None = None

    @property
    def terminal(self) -> FTFCmd:
        """惰性创建底层命令行对象（创建时需要 config.ftfpath 有效）"""
        if self._terminal is None:
            self._terminal = FTFCmd()
        return self._terminal

    @property
    def admin(self) -> bool:
        """当前是否已通过创始人密钥校验"""
        return isinstance(self._terminal, FTFAdminCmd)

    def authenticate(self, key: str) -> bool:
        """用《朝花夕拾协议》的宗旨校验创始人权限，与命令行进入方式相同"""
        if not key or not check_founder_key(key):
            return False
        with self._lock:
            self._terminal = FTFAdminCmd()
        return True

    def commands(self) -> list[str]:
        """返回当前权限下可用的、已封装为 API 的命令名列表"""
        terminal = self.terminal
        return sorted(
            name[3:]
            for name in dir(terminal)
            if name.startswith("do_")
            and callable(getattr(terminal, name))
            and name[3:] not in _NOT_API
        )

    def describe(self) -> list[CommandInfo]:
        """返回全部命令的名称、摘要、语法与完整帮助文档"""
        terminal = self.terminal
        infos: list[CommandInfo] = []
        for name in self.commands():
            doc = inspect.getdoc(getattr(terminal, f"do_{name}")) or ""
            lines = [line.strip() for line in doc.splitlines() if line.strip()]
            summary = lines[0] if lines else ""
            syntax = next(
                (line for line in lines if line.startswith("语法：")),
                "",
            )
            infos.append(
                CommandInfo(name=name, summary=summary, syntax=syntax, doc=doc)
            )
        return infos

    def help(self, name: str) -> str:
        """返回指定命令的帮助文档；命令不存在时返回空字符串"""
        handler = getattr(self.terminal, f"do_{name.lstrip('/')}", None)
        if handler is None or not callable(handler):
            return ""
        return inspect.getdoc(handler) or ""

    def run(
        self, command: str, *, answer: bool | None = None
    ) -> CommandResult:
        """执行一条完整的命令行指令，例如 run("find 愿望 in year 2025")"""
        command = command.strip()
        if command.startswith("/"):
            command = command[1:]
        if not command:
            return CommandResult(command="", ok=False, error="命令为空")
        name, _, args = command.partition(" ")
        return self.call(name, args, answer=answer)

    def call(
        self, name: str, *args: Any, answer: bool | None = None
    ) -> CommandResult:
        """按命令名调用，参数可传多个字符串或一个字符串/列表

        call("find", "愿望", "in", "year", "2025")
        call("find", "愿望 in year 2025")
        call("find", ["愿望", "in", "year", "2025"])

        answer 用于回答命令内部的人工确认：
            None  不回答，返回 needs_confirmation=True 交由调用方决定
            True  确认（相当于命令行输入 Y）
            False 取消（相当于命令行输入 N）
        """
        name = name.strip().lstrip("/")
        arg_text = self._join(args)
        line = f"{name} {arg_text}".strip()

        handler = getattr(self.terminal, f"do_{name}", None)
        if not name or handler is None or not callable(handler):
            return CommandResult(
                command=line, ok=False, error=f"没有此命令: {name}"
            )

        buffer = io.StringIO()
        with self._lock:
            try:
                with self._capture(buffer), self._choice_answers(answer):
                    handler(arg_text)
            except ConfirmationRequired as e:
                return CommandResult(
                    command=line,
                    ok=False,
                    output=self._clean(buffer.getvalue()),
                    needs_confirmation=True,
                    question=e.question,
                )
            except CommandLineExit:
                return CommandResult(
                    command=line,
                    ok=True,
                    output=self._clean(buffer.getvalue()),
                    exit_requested=True,
                )
            except AdminMode:
                return CommandResult(
                    command=line,
                    ok=False,
                    output=self._clean(buffer.getvalue()),
                    error="命令需要创始人权限，请先通过 authenticate() 解锁",
                )
            except Exception as e:
                return CommandResult(
                    command=line,
                    ok=False,
                    output=self._clean(buffer.getvalue()),
                    error=f"{type(e).__name__}: {e}",
                )

        return CommandResult(
            command=line, ok=True, output=self._clean(buffer.getvalue())
        )

    def find(self, *args: Any, answer: bool | None = None) -> CommandResult:
        """查找文档中的关键字词。语法：find <keywords> in <范围>"""
        return self.call("find", *args, answer=answer)

    def count(self, *args: Any, answer: bool | None = None) -> CommandResult:
        """统计年度总结中的新事物/重大事件或“时期”数量"""
        return self.call("count", *args, answer=answer)

    def check(self, *args: Any, answer: bool | None = None) -> CommandResult:
        """检查事件记录文档及年度总结的格式规范性。语法：check <years>"""
        return self.call("check", *args, answer=answer)

    def calculate(
        self, *args: Any, answer: bool | None = None
    ) -> CommandResult:
        """计算“时期”的强度"""
        return self.call("calculate", *args, answer=answer)

    def dellog(self, *args: Any, answer: bool | None = None) -> CommandResult:
        """删除日志文件（创始人命令）。语法：dellog <logname/*>"""
        return self.call("dellog", *args, answer=answer)

    def listfiles(
        self, *args: Any, answer: bool | None = None
    ) -> CommandResult:
        """列出指定年份的云端文件（创始人命令）。语法：listfiles <year>"""
        return self.call("listfiles", *args, answer=answer)

    def verify(self, *args: Any, answer: bool | None = None) -> CommandResult:
        """校验云端文件与本地文件是否一致（创始人命令）。语法：verify <year> [/reset]"""
        return self.call("verify", *args, answer=answer)

    def upload(self, *args: Any, answer: bool | None = None) -> CommandResult:
        """上传本地文档到云端（创始人命令）。语法：upload <year>/<month|annual_summary>"""
        return self.call("upload", *args, answer=answer)

    @staticmethod
    def _join(args: tuple[Any, ...]) -> str:
        if len(args) == 1 and isinstance(args[0], (list, tuple)):
            return " ".join(str(item) for item in args[0])
        return " ".join(str(item) for item in args)

    @staticmethod
    def _clean(text: str) -> str:
        return _ANSI_PATTERN.sub("", text).rstrip()

    @contextlib.contextmanager
    def _capture(self, buffer: io.StringIO) -> Iterator[None]:
        """同时捕获 print 与 cmd.Cmd 通过 self.stdout 写出的内容"""
        terminal = self.terminal
        original_stdout = terminal.stdout
        terminal.stdout = buffer
        try:
            with contextlib.redirect_stdout(buffer):
                yield
        finally:
            terminal.stdout = original_stdout

    @contextlib.contextmanager
    def _choice_answers(self, answer: bool | None) -> Iterator[None]:
        """接管命令内部的 choice() 确认

        answer=True  自动确认；answer=False 自动取消；
        answer=None  抛出 ConfirmationRequired，把确认交回调用方
        """
        original = ftf_cmd.choice

        if answer is None:

            def ask(
                choose="YN",
                text="Y/N",
                default=None,
                timeout=10,
                *,
                hide=False,
            ):
                raise ConfirmationRequired(text)

        else:
            value = 1 if answer else 2

            def ask(
                choose="YN",
                text="Y/N",
                default=None,
                timeout=10,
                *,
                hide=False,
            ):
                return value

        ftf_cmd.choice = ask
        try:
            yield
        finally:
            ftf_cmd.choice = original


# 模块级默认实例：默认只有公开命令权限，需 authenticate() 后才有创始人权限
terminal = TerminalAPI()


def run(command: str, *, answer: bool | None = None) -> CommandResult:
    """执行一条完整的命令行指令"""
    return terminal.run(command, answer=answer)


def authenticate(key: str) -> bool:
    """用《朝花夕拾协议》的宗旨校验创始人权限"""
    return terminal.authenticate(key)


def is_admin() -> bool:
    """当前是否已通过创始人密钥校验"""
    return terminal.admin


def commands() -> list[str]:
    """返回可用命令名列表"""
    return terminal.commands()


def describe() -> list[CommandInfo]:
    """返回全部命令的描述信息"""
    return terminal.describe()


def help_command(name: str) -> str:
    """返回指定命令的帮助文档（命令名可为 help 子命令，如 verify）"""
    return terminal.help(name)


def find(*args: Any, answer: bool | None = None) -> CommandResult:
    """查找文档中的关键字词"""
    return terminal.find(*args, answer=answer)


def count(*args: Any, answer: bool | None = None) -> CommandResult:
    """统计年度总结中的新事物/重大事件或“时期”数量"""
    return terminal.count(*args, answer=answer)


def check(*args: Any, answer: bool | None = None) -> CommandResult:
    """检查事件记录文档及年度总结的格式规范性"""
    return terminal.check(*args, answer=answer)


def calculate(*args: Any, answer: bool | None = None) -> CommandResult:
    """计算“时期”的强度"""
    return terminal.calculate(*args, answer=answer)


def dellog(*args: Any, answer: bool | None = None) -> CommandResult:
    """删除日志文件（创始人命令）"""
    return terminal.dellog(*args, answer=answer)


def listfiles(*args: Any, answer: bool | None = None) -> CommandResult:
    """列出指定年份的云端文件（创始人命令）"""
    return terminal.listfiles(*args, answer=answer)


def verify(*args: Any, answer: bool | None = None) -> CommandResult:
    """校验云端文件与本地文件是否一致（创始人命令）"""
    return terminal.verify(*args, answer=answer)


def upload(*args: Any, answer: bool | None = None) -> CommandResult:
    """上传本地文档到云端（创始人命令）"""
    return terminal.upload(*args, answer=answer)
