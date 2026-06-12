"""Qwen-OCR 解析适配器: 纯图片 PDF → SlideIR.

针对 NotebookLM 等导出的**纯图片 PDF**（每页一张光栅图、零矢量文字、零字体），
Docling 矢量解析读不到任何内容。此适配器改用阿里云百炼 Qwen-OCR 的
`advanced_recognition` 内置任务：逐页做文字行检测 + 识别，直接返回**以原图左上角
为原点的绝对像素坐标**（四顶点 location），无需反推 smart-resize 缩放。

坐标契约（来自百炼文档）：
- location = [x1,y1, x2,y2, x3,y3, x4,y4]，顶点顺序固定：左上→右上→右下→左下，
  原点 (0,0) 在原图左上角，单位为原图像素。
- 我们把四顶点折叠为轴对齐 BBox（min/max），与 SlideIR.BBox 的屏幕坐标系一致。

注意：`advanced_recognition` 的 ocr_options 只能经 DashScope 原生 SDK 传入，
OpenAI 兼容端点不支持该参数。
"""

from __future__ import annotations

import uuid
from pathlib import Path

from ppt_enhance.config import settings
from ppt_enhance.parser.pdf_renderer import get_page_size, render_pdf_pages
from ppt_enhance.schemas.slide_ir import BBox, ElementType, SlideElement, SlideIR, SlidePage

# advanced_recognition 任务的固定 Prompt（由模型内部使用，仅作占位说明）
_ADV_PROMPT = "定位所有的文字行，并且返回旋转矩形([cx, cy, width, height, angle])的坐标结果。"


def _ocr_page(local_path: str | Path, api_key: str, model: str, enable_rotate: bool = False) -> list[dict]:
    """对单页图调用 Qwen-OCR advanced_recognition，返回 words_info 列表。

    每个元素形如 {"text": "...", "location": [x1,y1,...,x4,y4],
    "rotate_rect": [cx,cy,w,h,angle]}。失败或无文字时返回 []。
    """
    import dashscope

    image_uri = f"file://{Path(local_path).resolve()}"
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "image": image_uri,
                    # 像素阈值：小于下限放大、大于上限缩小（保持原比例）
                    "min_pixels": 28 * 28 * 4,
                    "max_pixels": 28 * 28 * 8192,
                    "enable_rotate": enable_rotate,
                },
                {"text": _ADV_PROMPT},
            ],
        }
    ]
    response = dashscope.MultiModalConversation.call(
        api_key=api_key,
        model=model,
        messages=messages,
        ocr_options={"task": "advanced_recognition"},
    )

    if getattr(response, "status_code", 200) != 200:
        raise RuntimeError(
            f"Qwen-OCR 调用失败: status={getattr(response, 'status_code', '?')} "
            f"code={getattr(response, 'code', '?')} msg={getattr(response, 'message', '')}"
        )

    content = response["output"]["choices"][0]["message"]["content"]
    # content 是 list[dict]；advanced_recognition 结果在 ocr_result 字段
    for part in content:
        if isinstance(part, dict) and "ocr_result" in part:
            ocr_result = part["ocr_result"]
            if isinstance(ocr_result, dict) and "words_info" in ocr_result:
                return ocr_result["words_info"] or []
    return []


def _poly_to_bbox(location: list[float]) -> BBox:
    """四顶点 [x1,y1,...,x4,y4] → 轴对齐 BBox。"""
    xs = [float(location[i]) for i in range(0, 8, 2)]
    ys = [float(location[i]) for i in range(1, 8, 2)]
    return BBox(x0=min(xs), y0=min(ys), x1=max(xs), y1=max(ys))


def _classify(bbox: BBox, page_w: float, page_h: float, text: str, idx: int) -> ElementType:
    """基于位置/尺寸的轻量启发式分类。

    Qwen-OCR 不输出语义标签，故用规则推断：
    - 高度占比大且靠上 → 标题
    - 行首为项目符号/序号 → 列表
    - 贴近页面底部且字小 → 页脚
    - 其余 → 正文
    """
    h_ratio = bbox.height / page_h if page_h else 0.0
    y_center_ratio = (bbox.y0 + bbox.y1) / 2 / page_h if page_h else 0.0
    stripped = text.lstrip()

    bullet_starts = ("·", "•", "-", "*", "‣", "▪", "◦", "·")
    if stripped[:1] in bullet_starts or (stripped[:2].rstrip(".、)）").isdigit()):
        return ElementType.LIST

    if y_center_ratio > 0.9 and h_ratio < 0.05:
        return ElementType.FOOTER

    # 首个较大文本块视作标题
    if idx == 0 and h_ratio > 0.06 and y_center_ratio < 0.5:
        return ElementType.TITLE
    if h_ratio > 0.08 and y_center_ratio < 0.4:
        return ElementType.TITLE

    return ElementType.TEXT


def parse_with_qwen_ocr(
    pdf_path: str | Path,
    work_dir: str | Path,
    dpi: int = 150,
    api_key: str | None = None,
    model: str = "qwen-vl-ocr-latest",
    enable_rotate: bool = False,
    progress_cb=None,
) -> SlideIR:
    """使用 Qwen-OCR 将纯图片 PDF 解析为 SlideIR.

    适用于 NotebookLM 等导出的、无矢量文字层的图片型 PDF。
    每页一次 OCR VLM 调用，按页并行（DashScope 客户端线程安全）。
    progress_cb(done, total) 在每页 OCR 完成时回调（主线程）。
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    api_key = api_key or settings.openai_api_key
    if not api_key:
        raise RuntimeError(
            "Qwen-OCR 需要百炼 API Key。请在 .env 中设置 OPENAI_API_KEY（DashScope key）。"
        )

    pdf_path = Path(pdf_path)
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    render_dir = work_dir / "renders"
    page_images = render_pdf_pages(pdf_path, render_dir, dpi=dpi)
    page_sizes = get_page_size(pdf_path, dpi=dpi)

    page_nos = sorted(page_images)
    total = len(page_nos)

    def ocr_one(page_no):
        png = page_images[page_no]
        words_info = _ocr_page(png, api_key=api_key, model=model, enable_rotate=enable_rotate)
        return page_no, words_info

    # 并行 OCR：网络等待为主，6 路并发显著缩短墙钟时间
    ocr_results: dict[int, list[dict]] = {}
    workers = min(settings.vlm_concurrency, max(1, total))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = [ex.submit(ocr_one, pno) for pno in page_nos]
        for done_n, fut in enumerate(as_completed(futures), start=1):
            pno, words_info = fut.result()
            ocr_results[pno] = words_info
            if progress_cb:
                progress_cb(done_n, total)

    pages: list[SlidePage] = []
    all_texts: list[str] = []
    total_elems = 0

    for page_no in page_nos:  # 按页序构建
        png = page_images[page_no]
        w, h = page_sizes.get(page_no, (1920, 1080))
        words_info = ocr_results[page_no]

        elements: list[SlideElement] = []
        for idx, line in enumerate(words_info):
            text = (line.get("text") or "").strip()
            loc = line.get("location")
            if not text or not loc or len(loc) < 8:
                continue
            bbox = _poly_to_bbox(loc)
            if bbox.area <= 0:
                continue
            elem_type = _classify(bbox, w, h, text, idx)
            all_texts.append(text)
            elements.append(
                SlideElement(
                    id=f"ocr_{uuid.uuid4().hex[:8]}",
                    type=elem_type,
                    text=text,
                    bbox=bbox,
                    page_no=page_no,
                    font_size_hint=max(8.0, bbox.height * 0.75),
                    metadata={"rotate_rect": line.get("rotate_rect")},
                )
            )

        total_elems += len(elements)
        pages.append(
            SlidePage(
                page_no=page_no,
                width=w,
                height=h,
                render_path=str(png),
                elements=elements,
            )
        )

    from ppt_enhance.parser.docling_adapter import _extract_protected_terms

    return SlideIR(
        source_pdf=str(pdf_path.resolve()),
        parser="qwen-ocr",
        dpi=dpi,
        pages=pages,
        protected_terms=_extract_protected_terms(all_texts),
        metadata={"element_count": total_elems, "ocr_model": model},
    )
