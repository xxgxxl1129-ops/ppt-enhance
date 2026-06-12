"""OpenAI 兼容 LLM 客户端."""

from __future__ import annotations

import base64
import json
import threading
from pathlib import Path
from typing import Any

from openai import OpenAI

from ppt_enhance.config import settings


def _encode_image(image_path: str | Path) -> str:
    """本地图片 → data URI（base64），供视觉模型传入。"""
    p = Path(image_path)
    data = p.read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    suffix = p.suffix.lower().lstrip(".") or "png"
    mime = "jpeg" if suffix in ("jpg", "jpeg") else suffix
    return f"data:image/{mime};base64,{b64}"


class LLMClient:
    def __init__(
        self,
        api_key: str | None = None,
        base_url: str | None = None,
        model: str | None = None,
    ):
        self.api_key = api_key or settings.openai_api_key
        self.base_url = base_url or settings.openai_base_url
        self.model = model or settings.openai_model
        self._client: OpenAI | None = None
        self._client_lock = threading.Lock()

    @property
    def available(self) -> bool:
        return bool(self.api_key)

    @property
    def client(self) -> OpenAI:
        # 双重检查锁：多线程并行调用时只初始化一个底层 client
        if self._client is None:
            with self._client_lock:
                if self._client is None:
                    self._client = OpenAI(api_key=self.api_key, base_url=self.base_url)
        return self._client

    def chat_json(self, system: str, user: str, temperature: float = 0.2) -> dict[str, Any]:
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            temperature=temperature,
            response_format={"type": "json_object"},
        )
        content = response.choices[0].message.content or "{}"
        return json.loads(content)

    def vision_json(
        self,
        system: str,
        user_text: str,
        image_path: str | Path,
        model: str | None = None,
        temperature: float = 0.1,
    ) -> dict[str, Any]:
        """多模态调用：传入页面图 + 指令，要求模型返回 JSON。

        用于版面/形状识别（qwen-vl 系列）。失败时返回 {}。
        """
        vision_model = model or "qwen-vl-plus"
        data_uri = _encode_image(image_path)
        response = self.client.chat.completions.create(
            model=vision_model,
            messages=[
                {"role": "system", "content": system},
                {
                    "role": "user",
                    "content": [
                        {"type": "image_url", "image_url": {"url": data_uri}},
                        {"type": "text", "text": user_text},
                    ],
                },
            ],
            temperature=temperature,
        )
        content = response.choices[0].message.content or "{}"
        # qwen-vl 可能包 ```json ``` 代码块，剥掉
        content = content.strip()
        if content.startswith("```"):
            content = content.split("```", 2)[1]
            if content.startswith("json"):
                content = content[4:]
            content = content.strip().rstrip("`").strip()
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            # 尝试截取第一个 { 到最后一个 }
            s, e = content.find("{"), content.rfind("}")
            if s >= 0 and e > s:
                try:
                    return json.loads(content[s : e + 1])
                except json.JSONDecodeError:
                    return {}
            return {}
