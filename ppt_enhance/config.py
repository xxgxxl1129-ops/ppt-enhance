"""全局配置."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


@dataclass(frozen=True)
class Settings:
    openai_api_key: str = os.getenv("OPENAI_API_KEY", "")
    openai_base_url: str = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
    openai_model: str = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    max_correction_rounds: int = int(os.getenv("MAX_CORRECTION_ROUNDS", "3"))
    max_edit_ratio: float = float(os.getenv("MAX_EDIT_RATIO", "0.15"))
    default_dpi: int = int(os.getenv("DEFAULT_DPI", "150"))
    work_dir: Path = Path(os.getenv("WORK_DIR", ".ppt_enhance_cache"))
    # VLM 并发数：OCR 解析与大纲逆推按页并行的线程数。
    # 默认 6（与原 batch_outline.py 一致，对 DashScope 默认额度安全）；
    # key 并发额度高时可调大提速，过高会触发限流。
    vlm_concurrency: int = int(os.getenv("VLM_CONCURRENCY", "6"))


settings = Settings()
