"""
硅基流动（SiliconFlow）OpenAI 兼容客户端
只依赖 requests，不引入额外 SDK。用于把自然语言请求转换成语义化工具调用。
配置来自 config.ini 的 [AI] 段；环境变量 SILICONFLOW_API_KEY 优先于配置文件。
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass, field
from typing import Any

import requests

from libs.Common import config
from libs.Common.exceptions import AIError
from libs.Common.functions import log

DEFAULT_BASE_URL = "https://api.siliconflow.cn/v1"

# 这些模型支持 enable_thinking 开关；其它模型不发送该字段，避免被网关拒绝
_THINKING_MODELS = ("qwen3", "qwq", "glm-z1", "deepseek-r1", "hunyuan-a13b")


@dataclass(slots=True)
class ToolCall:
    """模型请求调用的一次工具"""

    id: str
    name: str
    arguments: str

    def parse_arguments(self) -> dict[str, Any]:
        """把 arguments 解析成字典；不是合法 JSON 对象时抛出 AIError"""
        try:
            value = json.loads(self.arguments or "{}")
        except json.JSONDecodeError as e:
            raise AIError(f"工具参数不是合法 JSON: {e}") from e
        if not isinstance(value, dict):
            raise AIError("工具参数必须是 JSON 对象")
        return value

    def to_message(self) -> dict[str, Any]:
        """转换回 OpenAI 消息格式，用于把校验错误反馈给模型"""
        return {
            "id": self.id,
            "type": "function",
            "function": {"name": self.name, "arguments": self.arguments},
        }


@dataclass(slots=True)
class ChatReply:
    """一次模型回复"""

    content: str = ""
    reasoning: str = ""
    tool_calls: list[ToolCall] = field(default_factory=list)
    finish_reason: str = ""
    usage: dict[str, Any] = field(default_factory=dict)
    model: str = ""

    @property
    def has_tool_calls(self) -> bool:
        return bool(self.tool_calls)


class AIClient:
    """OpenAI 兼容的对话客户端"""

    def __init__(
        self,
        api_key: str | None = None,
        base_url: str | None = None,
        model: str | None = None,
        timeout: float | None = None,
        max_tokens: int | None = None,
        temperature: float | None = None,
        enable_thinking: bool | None = None,
        retries: int = 2,
    ) -> None:
        self.api_key = (
            api_key
            or os.environ.get("SILICONFLOW_API_KEY", "")
            or config.ai_api_key
        ).strip()
        self.base_url = (
            base_url or config.ai_base_url or DEFAULT_BASE_URL
        ).rstrip("/")
        self.model = model or config.ai_model
        self.timeout = timeout if timeout is not None else config.ai_timeout
        self.max_tokens = (
            max_tokens if max_tokens is not None else config.ai_max_tokens
        )
        self.temperature = (
            temperature if temperature is not None else config.ai_temperature
        )
        self.enable_thinking = (
            enable_thinking
            if enable_thinking is not None
            else config.ai_enable_thinking
        )
        self.retries = max(0, retries)

    @property
    def configured(self) -> bool:
        """密钥是否已填写（排除 config.ini 里的占位符）"""
        return bool(self.api_key) and not self.api_key.startswith("{")

    def chat(
        self,
        messages: list[dict[str, Any]],
        tools: list[dict[str, Any]] | None = None,
        *,
        tool_choice: str | None = None,
        temperature: float | None = None,
        max_tokens: int | None = None,
    ) -> ChatReply:
        """发送一次非流式对话请求"""
        if not self.configured:
            raise AIError(
                "未配置硅基流动 API 密钥，请在 config.ini 的 [AI] 段填写 api_key"
            )

        payload: dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            "temperature": (
                self.temperature if temperature is None else temperature
            ),
            "max_tokens": (
                self.max_tokens if max_tokens is None else max_tokens
            ),
        }
        if tools:
            payload["tools"] = tools
            payload["tool_choice"] = tool_choice or "auto"
        if self._supports_thinking():
            payload["enable_thinking"] = self.enable_thinking

        data = self._post(payload)
        reply = self._parse_reply(data)
        self._log_reply(reply)
        return reply

    @staticmethod
    def _log_reply(reply: ChatReply) -> None:
        """把模型的输出写入日志，便于事后核对它到底说了什么"""
        if reply.reasoning:
            log(f"[AI 思考] {reply.reasoning}", "info", logfile_only=True)
        if reply.content:
            log(f"[AI 输出] {reply.content}", "info", logfile_only=True)
        for call in reply.tool_calls:
            log(
                f"[AI 工具调用] {call.name} {call.arguments}",
                "info",
                logfile_only=True,
            )
        usage = reply.usage or {}
        if usage:
            log(
                "AI 用量: prompt={0} completion={1} total={2}".format(
                    usage.get("prompt_tokens"),
                    usage.get("completion_tokens"),
                    usage.get("total_tokens"),
                ),
                "info",
                logfile_only=True,
            )

    def _supports_thinking(self) -> bool:
        name = self.model.lower()
        return any(key in name for key in _THINKING_MODELS)

    def _post(self, payload: dict[str, Any]) -> dict[str, Any]:
        """带退避重试的 POST；只记录状态码，不记录请求内容"""
        url = f"{self.base_url}/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        last_error = ""
        for attempt in range(self.retries + 1):
            try:
                response = requests.post(
                    url,
                    headers=headers,
                    json=payload,
                    timeout=self.timeout,
                )
            except requests.RequestException as e:
                last_error = f"请求失败: {e}"
                if attempt < self.retries:
                    log(
                        f"[AI 请求失败] 准备重试: {type(e).__name__}",
                        "warning",
                    )
                    time.sleep(min(2**attempt, 8))
                    continue
                raise AIError(last_error) from e

            if response.status_code == 200:
                try:
                    data = response.json()
                except ValueError as e:
                    raise AIError("AI 返回的不是合法 JSON") from e
                if not isinstance(data, dict):
                    raise AIError("AI 返回格式异常")
                return data

            body = response.text[:300].replace("\n", " ")
            if self.api_key:
                body = body.replace(self.api_key, "***")
            last_error = f"AI 服务返回 {response.status_code}: {body}"
            if response.status_code == 429 or response.status_code >= 500:
                if attempt < self.retries:
                    log(
                        f"[AI 请求失败] 服务暂时不可用（{response.status_code}），准备重试",
                        "warning",
                    )
                    time.sleep(min(2**attempt, 8))
                    continue
            raise AIError(last_error)

        raise AIError(last_error or "AI 服务调用失败")

    @staticmethod
    def _parse_reply(data: dict[str, Any]) -> ChatReply:
        choices = data.get("choices") or []
        if not choices:
            raise AIError("AI 没有返回任何候选结果")

        choice = choices[0] or {}
        message = choice.get("message") or {}
        tool_calls: list[ToolCall] = []
        for index, item in enumerate(message.get("tool_calls") or []):
            function = item.get("function") or {}
            arguments = function.get("arguments") or "{}"
            if not isinstance(arguments, str):
                arguments = json.dumps(arguments, ensure_ascii=False)
            tool_calls.append(
                ToolCall(
                    id=item.get("id") or f"call_{index}",
                    name=function.get("name") or "",
                    arguments=arguments,
                )
            )

        content = message.get("content") or ""
        reasoning = (
            message.get("reasoning_content") or message.get("reasoning") or ""
        )
        return ChatReply(
            content=(
                content.strip() if isinstance(content, str) else str(content)
            ),
            reasoning=(
                reasoning.strip()
                if isinstance(reasoning, str)
                else str(reasoning)
            ),
            tool_calls=tool_calls,
            finish_reason=choice.get("finish_reason") or "",
            usage=data.get("usage") or {},
            model=data.get("model") or "",
        )
