"""字体与文本框尺寸工具."""

from __future__ import annotations

import sys

from pptx.util import Pt


def default_cjk_font() -> str:
    """按运行平台返回一个该平台预装的中文字体名。

    python-pptx 写入的字体名只是「首选」标签，渲染端（PowerPoint /
    LibreOffice）据此查找字体，找不到才回退。CJK 字体三平台命名各异，
    故按当前平台选该平台必有的中文字体，保证「在哪生成、在那预览」一致，
    避免写死 Windows 专属字体导致其他平台被替换、排版错位。
    """
    if sys.platform == "darwin":
        return "PingFang SC"
    if sys.platform.startswith("win"):
        return "Microsoft YaHei"
    return "Noto Sans CJK SC"  # Linux 常见；缺失时 LO 仍会找等价 CJK


def estimate_font_size(bbox_height: float, line_count: int = 1, min_pt: float = 8.0) -> float:
    """根据 bbox 高度估算字号（磅）."""
    if line_count <= 1:
        pt = bbox_height * 0.75 * 0.75  # px → pt 近似
    else:
        pt = (bbox_height / line_count) * 0.75 * 0.75
    return max(min_pt, pt)


def pt_to_ppt(pt: float) -> Pt:
    return Pt(pt)
