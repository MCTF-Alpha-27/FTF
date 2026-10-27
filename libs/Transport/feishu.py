"""
飞书自建应用传输层
"""
import json
import os
import queue
import threading
import time

import lark_oapi as lark
import requests
from lark_oapi.api.im.v1 import (
    CreateFileRequest,
    CreateFileRequestBody,
    CreateImageRequest,
    CreateImageRequestBody,
    CreateMessageRequest,
    CreateMessageRequestBody,
    DeleteMessageRequest,
    P2ImMessageReceiveV1,
)
from lark_oapi.event.callback.model.p2_card_action_trigger import (
    CallBackCard,
    CallBackToast,
    P2CardActionTrigger,
    P2CardActionTriggerResponse,
)

from ..Common import config
from ..Common.functions import log
from .remote import RemoteTerminal

MAX_SEEN = 200
MAX_FILE_SIZE = 30 * 1024 * 1024
MAX_IMAGE_SIZE = 10 * 1024 * 1024

# 密钥输入卡片中密码框的字段名
AUTH_INPUT_NAME = "founder_key"

def _auth_input_card() -> dict:
    """带密钥输入框的创始人密钥卡片，密钥不会作为聊天消息发送

    注意：不能用 input_type=password，密码框会禁用输入法导致无法输入中文
    提交后卡片会立即被结果卡片替换，输入内容不会留在聊天记录中
    """
    return {
        "config": {"wide_screen_mode": True, "update_multi": True},
        "header": {
            "template": "blue",
            "title": {"tag": "plain_text", "content": "创始人权限校验"},
        },
        "elements": [
            {
                "tag": "div",
                "text": {
                    "tag": "plain_text",
                    "content": "请输入《朝花夕拾协议》的宗旨。密钥在卡片内提交，不会出现在聊天记录中，提交后卡片会被替换",
                },
            },
            {
                "tag": "form",
                "name": "founder_auth_form",
                "elements": [
                    {
                        "tag": "input",
                        "name": AUTH_INPUT_NAME,
                        "input_type": "text",
                        "required": True,
                        "max_length": 200,
                        "label": {"tag": "plain_text", "content": "密钥"},
                        "label_position": "top",
                        "placeholder": {"tag": "plain_text", "content": "请输入密钥"},
                    },
                    {
                        "tag": "button",
                        "action_type": "form_submit",
                        "name": "founder_auth_submit",
                        "text": {"tag": "plain_text", "content": "提交"},
                        "type": "primary",
                    },
                ],
            },
        ],
    }

def _auth_result_card(ok: bool) -> dict:
    """校验结束后替换原卡片，输入框与密钥一并从卡片中消失"""
    if ok:
        title, template = "创始人权限已解锁", "green"
        text = "密钥校验通过，现在可以执行创始人命令"
    else:
        title, template = "密钥校验失败", "red"
        text = "密钥错误，再次执行创始人命令可获取新的输入卡片"
    return {
        "config": {"wide_screen_mode": True, "update_multi": True},
        "header": {
            "template": template,
            "title": {"tag": "plain_text", "content": title},
        },
        "elements": [{"tag": "div", "text": {"tag": "plain_text", "content": text}}],
    }

class FeishuTransport(RemoteTerminal):
    """把飞书长连接事件翻译成指令，把回复翻译成飞书消息"""
    
    def __init__(
        self,
        app_id: str,
        app_secret: str,
        allowed_open_ids: list[str] | None = None,
        encrypt_key: str = "",
        verification_token: str = "",
    ) -> None:
        self.app_id = app_id
        self.app_secret = app_secret
        self.allowed = {i.strip() for i in (allowed_open_ids or []) if i.strip()}
        self.client = lark.Client.builder().app_id(app_id).app_secret(app_secret).build()
        self.inbox: queue.Queue = queue.Queue()
        self._chat_id: str = ""
        self._seen: list[str] = []
        self._seen_set: set[str] = set()
        self._stopped = False
        self._handler = (
            lark.EventDispatcherHandler.builder(encrypt_key, verification_token)
            .register_p2_im_message_receive_v1(self._on_message)
            .register_p2_card_action_trigger(self._on_card_action)
            .build()
        )
        self._ws = lark.ws.Client(
            app_id, app_secret, event_handler=self._handler, log_level=lark.LogLevel.INFO
        )

    def conversation_key(self) -> str:
        """同一飞书会话共享 AI 上下文"""
        return self._chat_id

    def _remember(self, message_id: str) -> None:
        self._seen.append(message_id)
        self._seen_set.add(message_id)
        if len(self._seen) > MAX_SEEN:
            self._seen_set.discard(self._seen.pop(0))

    def _on_message(self, data: P2ImMessageReceiveV1) -> None:
        try:
            if self._stopped:
                return
            event = data.event
            message = event.message
            sender = event.sender
            if sender is None or sender.sender_type != "user":
                return
            open_id = (sender.sender_id.open_id if sender.sender_id else "") or ""
            log(
                f"收到飞书消息: open_id={open_id}, chat_id={message.chat_id}, 类型={message.message_type}",
                "info",
                logfile_only=True,
            )
            if self.allowed and open_id not in self.allowed:
                log(f"忽略未授权飞书用户的消息: {open_id}", "warning", logfile_only=True)
                return
            if message.message_id in self._seen_set:
                return
            self._remember(message.message_id)
            if message.message_type != "text":
                self.inbox.put((open_id, message.chat_id, "<unsupported>"))
                return
            text = (json.loads(message.content or "{}").get("text") or "").strip()
            # 群里 @机器人 会留下 @_user_1 这类占位符
            text = " ".join(part for part in text.split() if not part.startswith("@_user_"))
            if text:
                self.inbox.put((open_id, message.chat_id, text))
        except Exception as e:
            log(f"解析飞书事件失败: {e}", "error")

    def _on_card_action(self, data: P2CardActionTrigger) -> P2CardActionTriggerResponse:
        """处理密钥卡片提交；回传的新卡片会覆盖输入框"""
        try:
            event = data.event
            operator = event.operator if event else None
            open_id = (operator.open_id if operator else "") or ""
            action = event.action if event else None
            form = (action.form_value if action and action.form_value else {}) or {}
            log(f"收到飞书卡片回调: open_id={open_id}", "info", logfile_only=True)

            if self.allowed and open_id not in self.allowed:
                log(f"忽略未授权飞书用户的卡片回调: {open_id}", "warning", logfile_only=True)
                return self._card_response(False, "无权操作")

            key = form.get(AUTH_INPUT_NAME, "") or ""
            if not key:
                return self._card_response(False, "请输入密钥")

            log("收到创始人密钥校验请求（卡片输入）", "info")
            if self.authenticate(key):
                log("创始人密钥校验通过", "info")
                self._auth_card_id = None
                return self._card_response(True, "创始人权限已解锁", _auth_result_card(True))
            log("创始人密钥校验失败", "warning")
            context = event.context if event else None
            if context and context.open_message_id:
                # 记住这张卡片，下次再尝试创始人命令时撤回它
                self._auth_card_id = context.open_message_id
            return self._card_response(False, "密钥错误", _auth_result_card(False))
        except Exception as e:
            log(f"处理飞书卡片回调失败: {e}", "error")
            return self._card_response(False, f"处理失败: {e}")

    @staticmethod
    def _card_response(
        ok: bool, message: str, card: dict | None = None
    ) -> P2CardActionTriggerResponse:
        response = P2CardActionTriggerResponse()
        response.toast = CallBackToast(
            {"type": "success" if ok else "error", "content": message}
        )
        if card is not None:
            response.card = CallBackCard({"type": "raw", "data": card})
        return response

    def send_auth_card(self) -> None:
        self._auth_card_id = self._send_card(_auth_input_card())

    def cancel_auth_card(self) -> None:
        """撤回上一次发送的密钥卡片"""
        message_id = self._auth_card_id
        if not message_id:
            return
        self._auth_card_id = None
        request = DeleteMessageRequest.builder().message_id(message_id).build()
        response = self.client.im.v1.message.delete(request)
        if not response.success():
            log(f"撤回密钥卡片失败（code={response.code}）: {response.msg}", "warning")
        else:
            log("已撤回上一个密钥卡片", "info", logfile_only=True)

    def reply(self, text: str) -> None:
        if not self._chat_id:
            return
        body = (
            CreateMessageRequestBody.builder()
            .receive_id(self._chat_id)
            .msg_type("text")
            .content(json.dumps({"text": text}, ensure_ascii=False))
            .build()
        )
        request = (
            CreateMessageRequest.builder()
            .receive_id_type("chat_id")
            .request_body(body)
            .build()
        )
        response = self.client.im.v1.message.create(request)
        if not response.success():
            raise Exception(f"飞书发送消息失败（code={response.code}）: {response.msg}")

    def _send_card(self, card: dict) -> str | None:
        if not self._chat_id:
            return None
        body = (
            CreateMessageRequestBody.builder()
            .receive_id(self._chat_id)
            .msg_type("interactive")
            .content(json.dumps(card, ensure_ascii=False))
            .build()
        )
        request = (
            CreateMessageRequest.builder()
            .receive_id_type("chat_id")
            .request_body(body)
            .build()
        )
        response = self.client.im.v1.message.create(request)
        if not response.success():
            raise Exception(f"飞书发送卡片失败（code={response.code}）: {response.msg}")
        return response.data.message_id if response.data else None

    def send_image(self, path: str) -> None:
        size = os.path.getsize(path)
        if size > MAX_IMAGE_SIZE:
            raise Exception(f"图片超过 10MB（{size} 字节），飞书无法发送")
        with open(path, "rb") as f:
            request = (
                CreateImageRequest.builder()
                .request_body(
                    CreateImageRequestBody.builder().image_type("message").image(f).build()
                )
                .build()
            )
            response = self.client.im.v1.image.create(request)
        if not response.success():
            raise Exception(f"飞书上传图片失败（code={response.code}）: {response.msg}")
        self._send_media("image", {"image_key": response.data.image_key})

    def send_file(self, path: str) -> None:
        size = os.path.getsize(path)
        if size > MAX_FILE_SIZE:
            raise Exception(f"文件超过 30MB（{size} 字节），飞书无法发送")
        with open(path, "rb") as f:
            request = (
                CreateFileRequest.builder()
                .request_body(
                    CreateFileRequestBody.builder()
                    .file_type("stream")
                    .file_name(os.path.basename(path))
                    .file(f)
                    .build()
                )
                .build()
            )
            response = self.client.im.v1.file.create(request)
        if not response.success():
            raise Exception(f"飞书上传文件失败（code={response.code}）: {response.msg}")
        self._send_media("file", {"file_key": response.data.file_key})

    def _send_media(self, msg_type: str, content: dict) -> None:
        if not self._chat_id:
            return
        body = (
            CreateMessageRequestBody.builder()
            .receive_id(self._chat_id)
            .msg_type(msg_type)
            .content(json.dumps(content))
            .build()
        )
        request = (
            CreateMessageRequest.builder()
            .receive_id_type("chat_id")
            .request_body(body)
            .build()
        )
        response = self.client.im.v1.message.create(request)
        if not response.success():
            raise Exception(f"飞书发送{msg_type}失败（code={response.code}）: {response.msg}")

    def check_credentials(self) -> tuple[bool, str]:
        """在建立长连接前先校验 App ID / App Secret，避免只看到晦涩的握手错误"""
        try:
            response = requests.post(
                "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
                json={"app_id": self.app_id, "app_secret": self.app_secret},
                timeout=10,
            )
            result = response.json()
        except Exception as e:
            return False, f"无法连接飞书开放平台: {e}"
        if result.get("code") != 0:
            return False, f"飞书凭据校验失败（code={result.get('code')}）: {result.get('msg')}"
        return True, "飞书凭据校验通过"

    def _ws_loop(self) -> None:
        try:
            self._ws.start()
        except Exception as e:
            self._stopped = True
            log(f"飞书长连接异常: {e}", "error")
            log("请确认开放平台已开启“使用长连接接收事件”、权限已发布、应用版本已上线", "error")
            self.inbox.put(None)

    def _work(self) -> str:
        # 指令统一在主线程顺序处理，避免与长连接线程并发操作界面和设备
        while not self._stopped:
            try:
                item = self.inbox.get(timeout=1)
            except queue.Empty:
                continue
            if item is None:
                return "error"
            open_id, chat_id, text = item
            self._chat_id = chat_id
            if text == "<unsupported>":
                self.notify("暂不支持该消息类型，请发送文本指令")
                continue
            if text == "/whoami":
                self.notify(f"open_id: {open_id}\nchat_id: {chat_id}")
                continue
            result = self.handle(text)
            if result == "exit":
                self._stopped = True
                time.sleep(0.5)
                # ws 在后台线程里阻塞且没有停止接口，直接结束进程最干净
                os._exit(0)
        return "stopped"

    def run(self) -> str:
        if not self.allowed:
            log("未配置飞书 allowed_open_ids，当前接受任意用户发来的指令（可用 /whoami 查看自己的 open_id）", "warning")
        ok, message = self.check_credentials()
        if not ok:
            log(message, "error")
            return "error"
        log(message, "info")
        log("正在建立飞书长连接", "info")
        threading.Thread(target=self._ws_loop, name="ftf-feishu-ws", daemon=True).start()
        return self._work()

def start() -> int:
    if not config.feishu_app_id or config.feishu_app_id.startswith("{"):
        log("未配置飞书 App ID，请在 config.ini 的 [Feishu] 段填写 app_id", "error")
        return 1
    if not config.feishu_app_secret or config.feishu_app_secret.startswith("{"):
        log("未配置飞书 App Secret，请在 config.ini 的 [Feishu] 段填写 app_secret", "error")
        return 1

    log("监听终端启动中", "info")
    if config.feishu_allowed_open_ids:
        log(f"已配置指令白名单，共{len(config.feishu_allowed_open_ids)}个 open_id", "info")
    else:
        log("未配置指令白名单，任何能给机器人发消息的人都可以执行指令", "warning")
        log("可先发送 /whoami 查看自己的 open_id，再填入 config.ini 的 allowed_open_ids", "warning")
    if config.ai_enabled:
        log(f"AI 语义解析已启用，模型 {config.ai_model}", "info")
    else:
        log("AI 语义解析未启用，仅响应以 / 开头的指令", "info")
    log("正在连接到飞书", "info")

    transport = FeishuTransport(
        config.feishu_app_id,
        config.feishu_app_secret,
        config.feishu_allowed_open_ids,
        config.feishu_encrypt_key,
        config.feishu_verification_token,
    )
    try:
        result = transport.run()
    except KeyboardInterrupt:
        log("监听终端已停止", "info")
        return 0
    except Exception as e:
        log(f"监听终端异常退出: {e}", "error")
        log("详细信息已写入日志", "exception")
        return 1
    return 1 if result == "error" else 0
