"""版面逆推（坐标锚定版）：OCR文字框 + VLM逻辑分组 → SlideOutline。

核心：坐标全部来自 OCR（准），VLM 只做「把哪些文字行归为同一张卡片/节点 + 谁连谁」
的逻辑判断。节点真实位置 = 其成员 OCR 框的外接框 → 重建版面与原图一致。
"""

from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np

from ppt_enhance.agents.llm_client import LLMClient
from ppt_enhance.schemas.outline import (
    Connection,
    FormulaImage,
    KeepRegion,
    LayoutType,
    OutlineNode,
    SlideOutline,
    StyleSpec,
)
from ppt_enhance.schemas.slide_ir import SlideElement

VISION_SYSTEM = "你是PPT版面结构分析专家。给定页面图和已识别的文字行(含id和坐标)，把文字行按逻辑归组为节点，并判断布局类型与连接关系。只输出JSON。"

VISION_PROMPT = """这是一张幻灯片(宽{w}px 高{h}px)。下面是已用OCR识别出的文字行，每行带唯一id和像素坐标(x0,y0,x1,y1)：

{lines}

请完成版面逆推：
1. 判断布局类型：title_cover/process_flow/timeline/comparison/list/table/tree/image_content/other
2. 把上述文字行按逻辑归组为「节点」(如一张卡片含 小标题+作者+正文 多行)。
3. 判断节点间的连接关系(流程箭头)。

输出JSON：
{{
  "layout_type": "...",
  "title_id": "页面主标题的id(无则空)",
  "subtitle_id": "副标题id(无则空)",
  "footer_id": "底部说明条id(无则空)",
  "nodes": [
    {{"order":1, "branch":0, "role":"节点角色",
      "heading_id":"小标题行id", "subtext_id":"次要文字行id",
      "body_ids":["正文行id1","正文行id2"]}}
  ],
  "connections": [{{"from":1,"to":2}}],
  "visual_regions": [
    {{"kind":"图表类型(曲线图/柱状图/雷达图/矩阵/插画等)", "bbox":[x0,y0,x1,y1], "desc":"这块画的什么"}}
  ],
  "design": {{
    "card_style": "卡片视觉风格: solid(实色填充)/translucent(半透明)/outline(纯描边) 三选一",
    "card_has_shadow": true,
    "card_radius": "圆角程度: 0(方角) 到 0.5(胶囊) 的小数",
    "header_band": "卡片顶部是否有强调色横条标签 true/false",
    "arrow_style": "箭头风格: solid(细线)/thick(粗线)/gradient(渐变) 三选一",
    "title_emphasis": "标题是否特大特粗 true/false"
  }}
}}

说明：
- branch：分叉流程的分支号，主干为0；步骤2后分叉到3、4，则3、4的branch为1、2。
- 每个id只能归到一个地方。标题/副标题/footer单独拎出，不要放进nodes。
- visual_regions：**关键**——只框出"重画会失真、必须保留原始像素"的真正复杂图形（柱状图/雷达图/曲线图/数据矩阵/精致插画）。
  普通的卡片、箭头、流程框、分隔线、图标**不算**，不要框（它们会用原生形状重画）。
  bbox 给图形所在的大致像素矩形即可，不必精确。这一页若没有此类复杂图形，给空数组 []。
- design：观察这页的实际视觉设计语言，如实描述（不同PPT设计不同，请据图判断）。
只输出JSON。"""


def _sample_theme_colors(image_path: str | Path) -> dict[str, str]:
    bgr = cv2.imread(str(image_path))
    if bgr is None:
        return {"bg_color": "#0E2A3F", "accent_color": "#27A68B", "text_color": "#FFFFFF"}
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    pixels = rgb.reshape(-1, 3).astype(np.int32)
    h, w = rgb.shape[:2]
    band = max(2, int(min(h, w) * 0.04))
    edges = np.concatenate([
        rgb[:band, :, :].reshape(-1, 3), rgb[-band:, :, :].reshape(-1, 3),
        rgb[:, :band, :].reshape(-1, 3), rgb[:, -band:, :].reshape(-1, 3),
    ]).astype(np.int32)
    bg = tuple(int(v) for v in np.median(edges, axis=0))
    hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV).reshape(-1, 3).astype(np.float32)
    score = (hsv[:, 1] / 255.0) * (hsv[:, 2] / 255.0)
    idx = np.argsort(score)[-max(1, len(score) // 50):]
    cand = pixels[idx]
    far = cand[np.linalg.norm(cand - np.array(bg), axis=1) > 50]
    if len(far) == 0:
        far = cand
    cq = (far // 16 * 16).astype(np.int32)
    ck = cq[:, 0] * 65536 + cq[:, 1] * 256 + cq[:, 2]
    cv_, cc_ = np.unique(ck, return_counts=True)
    ak = int(cv_[cc_.argmax()])
    acc = ((ak // 65536) % 256, (ak // 256) % 256, ak % 256)
    bg_lum = 0.299 * bg[0] + 0.587 * bg[1] + 0.114 * bg[2]
    text = (255, 255, 255) if bg_lum < 128 else (20, 20, 20)
    hx = lambda c: f"#{int(c[0]):02X}{int(c[1]):02X}{int(c[2]):02X}"
    return {"bg_color": hx(bg), "accent_color": hx(acc), "text_color": hx(text)}


def _union_bbox(elems: list[SlideElement]) -> list[float]:
    if not elems:
        return []
    x0 = min(e.bbox.x0 for e in elems)
    y0 = min(e.bbox.y0 for e in elems)
    x1 = max(e.bbox.x1 for e in elems)
    y1 = max(e.bbox.y1 for e in elems)
    return [x0, y0, x1, y1]


def detect_keep_regions(image_path, elements, bg_hex, vlm_regions, page_w, page_h):
    """VLM 语义门控 + OpenCV 几何定位的保留区检测。

    关键分工（基于实测：VLM 能准确识别"有什么图"，但给不准像素坐标）：
      - VLM（vlm_regions 非空）：语义判断本页确实含需保留的复杂图形 → 决定"要不要找"。
      - OpenCV：在全页实际检测"非背景非文字"的图形连通块 → 决定"在哪、多大"（真实坐标）。
    VLM 的 bbox 仅用于排序参考，不直接采用。
    """
    if not vlm_regions:
        return []
    bgr = cv2.imread(str(image_path))
    if bgr is None:
        return []
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB).astype(np.int32)
    h, w = rgb.shape[:2]
    try:
        bg = np.array([int(bg_hex[i:i + 2], 16) for i in (1, 3, 5)])
    except (ValueError, IndexError):
        bg = np.array([14, 42, 63])
    fg_full = (np.linalg.norm(rgb - bg, axis=2) > 45).astype(np.uint8)
    fg_all_density = float(fg_full.mean())
    # 抹掉文字（连同 padding），并记录文字覆盖掩膜
    fg = fg_full.copy()
    text_mask = np.zeros((h, w), np.uint8)
    for e in elements:
        b = e.bbox
        fg[max(0, int(b.y0) - 10):min(h, int(b.y1) + 10),
           max(0, int(b.x0) - 10):min(w, int(b.x1) + 10)] = 0
        text_mask[max(0, int(b.y0)):min(h, int(b.y1)),
                  max(0, int(b.x0)):min(w, int(b.x1))] = 1

    # 合并图形碎片成块
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (29, 29))
    closed = cv2.morphologyEx(fg, cv2.MORPH_CLOSE, k)
    n, _, stats, _ = cv2.connectedComponentsWithStats(closed, connectivity=8)
    cands = []
    for i in range(1, n):
        x, y, ww, hh, area = stats[i]
        frac = (ww * hh) / (w * h)
        if frac < 0.04 or frac > 0.9:
            continue
        if ww < w * 0.08 or hh < h * 0.08:
            continue
        sub_fg = fg[y:y + hh, x:x + ww]
        # 逐块判据（替代页级门控，更稳）：块内图形密度足够 且 文字覆盖低。
        # 真图表(曲线/柱/漏斗)：图形密度高、文字稀疏 → 保留；
        # 文字版卡片框：抹字后密度低 或 文字覆盖高 → 排除。
        graphic_density = float(sub_fg.mean())
        text_cov = float(text_mask[y:y + hh, x:x + ww].mean())
        if graphic_density < 0.06:
            continue
        if text_cov > 0.35:
            continue
        # 综合得分：图形密度高、文字少者优先
        score = area * graphic_density * (1.0 - text_cov)
        cands.append((score, [float(x), float(y), float(x + ww), float(y + hh)]))
    cands.sort(key=lambda c: -c[0])
    n_keep = min(len(cands), max(1, len(vlm_regions)), 3)
    return [c[1] for c in cands[:n_keep]]


def _sample_region_fill(image_path, bbox_px) -> str | None:
    """采样某区域的填充底色（取该区域非文字像素的中位色，即卡片背景）。"""
    if not bbox_px:
        return None
    bgr = cv2.imread(str(image_path))
    if bgr is None:
        return None
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    h, w = rgb.shape[:2]
    x0, y0, x1, y1 = [int(v) for v in bbox_px]
    # 略微外扩到卡片边缘区
    mx, my = int((x1 - x0) * 0.15), int((y1 - y0) * 0.15)
    x0, y0 = max(0, x0 - mx), max(0, y0 - my)
    x1, y1 = min(w, x1 + mx), min(h, y1 + my)
    if x1 <= x0 or y1 <= y0:
        return None
    patch = rgb[y0:y1, x0:x1].reshape(-1, 3).astype(np.int32)
    # 取最频繁色（量化）作为卡片底色——文字是少数高对比像素，会被多数底色淹没
    q = (patch // 16 * 16)
    keys = q[:, 0] * 65536 + q[:, 1] * 256 + q[:, 2]
    vals, counts = np.unique(keys, return_counts=True)
    k = int(vals[counts.argmax()])
    c = ((k // 65536) % 256, (k // 256) % 256, k % 256)
    return f"#{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def _build_style(design: dict, colors: dict, card_fill: str | None, fill_differs: bool = False) -> StyleSpec:
    """VLM 设计判断(分类) + 采样色(精确) + 像素证据(纠正) → 自适应样式规格。"""
    spec = StyleSpec()
    cs = design.get("card_style", "translucent")
    if cs in ("solid", "translucent", "outline"):
        spec.card_style = cs
    # 像素证据纠正：卡片区底色明显不同于背景 → 不是纯描边，用半透明实底（更耐看）
    if fill_differs and spec.card_style == "outline":
        spec.card_style = "translucent"
    # 反向：VLM 说实色但采样显示与背景几乎一致 → 降级描边，避免一片死板
    if not fill_differs and spec.card_style == "solid":
        spec.card_style = "translucent"
    spec.card_shadow = bool(design.get("card_has_shadow", True))
    try:
        spec.card_radius = max(0.0, min(0.5, float(design.get("card_radius", 0.12))))
    except (TypeError, ValueError):
        pass
    spec.header_band = bool(design.get("header_band", False))
    as_ = design.get("arrow_style", "solid")
    if as_ in ("solid", "thick", "gradient"):
        spec.arrow_style = as_
        spec.arrow_width = {"solid": 2.5, "thick": 5.0, "gradient": 5.0}[as_]
    if design.get("title_emphasis"):
        spec.title_size = 34.0
    spec.arrow_color = colors["accent_color"]
    spec.card_border = colors["accent_color"]
    spec.header_band_color = colors["accent_color"]
    # 卡片填充：采样到则用采样色，否则用背景提亮
    spec.card_fill = card_fill or colors["bg_color"]
    spec.card_opacity = 0.78
    return spec


def extract_outline(
    image_path: str | Path,
    page_no: int,
    width: float,
    height: float,
    elements: list[SlideElement],
    llm: LLMClient | None = None,
    vision_model: str = "qwen3-vl-plus",
) -> SlideOutline:
    """OCR坐标锚定的版面逆推。elements 为该页 OCR 文字元素(带 id 和准确 bbox)。"""
    llm = llm or LLMClient()
    colors = _sample_theme_colors(image_path)
    by_id = {e.id: e for e in elements}

    base = SlideOutline(page_no=page_no, page_width=width, page_height=height,
                        source_image_path=str(image_path), **colors)
    if not llm.available or not elements:
        return base

    lines = "\n".join(
        f"  {e.id}: ({e.bbox.x0:.0f},{e.bbox.y0:.0f},{e.bbox.x1:.0f},{e.bbox.y1:.0f}) \"{e.final_text}\""
        for e in elements
    )
    prompt = VISION_PROMPT.format(w=int(width), h=int(height), lines=lines)
    try:
        data = llm.vision_json(VISION_SYSTEM, prompt, image_path, model=vision_model)
    except Exception:
        return base
    if not data:
        return base

    try:
        layout = LayoutType(data.get("layout_type", "other"))
    except ValueError:
        layout = LayoutType.OTHER

    def text_of(eid):
        e = by_id.get(eid)
        return e.final_text if e else ""

    def bbox_of_ids(ids):
        elems = [by_id[i] for i in ids if i in by_id]
        return _union_bbox(elems)

    from ppt_enhance.builder.formula_render import is_formula_line, clean_inline_latex
    from ppt_enhance.schemas.outline import BodyLine

    nodes = []
    for i, n in enumerate(data.get("nodes", [])):
        body_ids = [bid for bid in n.get("body_ids", []) if bid in by_id]
        member_ids = [x for x in [n.get("heading_id"), n.get("subtext_id"), *body_ids] if x and x in by_id]
        # 逐行分类：公式行 → kind=formula（后续转 OMML）；其余 → 清洗内联 LaTeX 的纯文字
        body_lines = []
        for bid in body_ids:
            t = text_of(bid).strip()
            if not t:
                continue
            if is_formula_line(t):
                body_lines.append(BodyLine(text=t, kind="formula"))
            else:
                body_lines.append(BodyLine(text=clean_inline_latex(t), kind="text"))
        plain = " ".join(bl.text for bl in body_lines if bl.kind == "text").strip()
        nodes.append(OutlineNode(
            role=n.get("role", "node"),
            heading=clean_inline_latex(text_of(n.get("heading_id"))),
            subtext=text_of(n.get("subtext_id")),
            body=plain,
            body_lines=body_lines,
            order=int(n.get("order", i + 1)),
            branch=int(n.get("branch", 0)),
            element_ids=member_ids,
            bbox_px=bbox_of_ids(member_ids),
        ))

    connections = [
        Connection(from_order=int(c["from"]), to_order=int(c["to"]))
        for c in data.get("connections", [])
        if c.get("from") and c.get("to")
    ]

    base.layout_type = layout

    # 标题：VLM 选定后，合并紧邻的换行续行（多行标题常被只选一行）
    title_id = data.get("title_id")
    subtitle_id = data.get("subtitle_id")
    used_ids = set()
    for n in data.get("nodes", []):
        for k in ("heading_id", "subtext_id"):
            if n.get(k):
                used_ids.add(n[k])
        used_ids.update(n.get("body_ids", []) or [])
    if data.get("footer_id"):
        used_ids.add(data["footer_id"])
    # 注意：subtitle_id 暂不计入 used_ids。VLM 常把「跨行标题的下半句」
    # 误判为副标题；下面的续行合并若发现它在几何上紧贴标题（同 x 起点、
    # 紧邻、字高相近），应优先并回标题，再从副标题剔除。
    title_ids = []
    merged_subtitle = False
    if title_id and title_id in by_id:
        title_ids = [title_id]
        anchor = by_id[title_id].bbox
        cur_bottom = anchor.y1
        cur_x = anchor.x0
        cur_h = anchor.y1 - anchor.y0
        # 向下找续行：x起点接近、紧邻、字高接近、未被占用
        for e in sorted(elements, key=lambda e: e.bbox.y0):
            if e.id in title_ids or e.id in used_ids:
                continue
            b = e.bbox
            gap = b.y0 - cur_bottom
            if (abs(b.x0 - cur_x) < cur_h * 1.2 and -cur_h * 0.3 <= gap <= cur_h * 0.9
                    and abs((b.y1 - b.y0) - cur_h) < cur_h * 0.5):
                title_ids.append(e.id)
                cur_bottom = b.y1
                if e.id == subtitle_id:
                    merged_subtitle = True

    base.title = " ".join(text_of(i) for i in title_ids).strip()
    base.title_bbox_px = bbox_of_ids(title_ids) if title_ids else []
    # 续行已并回标题的副标题不再重复显示
    if merged_subtitle:
        subtitle_id = None
    base.subtitle = text_of(subtitle_id) if subtitle_id else ""
    base.subtitle_bbox_px = bbox_of_ids([subtitle_id]) if subtitle_id else []
    base.footer_note = text_of(data.get("footer_id"))
    base.footer_bbox_px = bbox_of_ids([data["footer_id"]]) if data.get("footer_id") else []
    base.nodes = nodes
    base.connections = connections

    # 自适应样式：从第一个节点区域采样卡片底色 + VLM 设计判断
    card_fill = None
    fill_differs = False
    if nodes and nodes[0].bbox_px:
        card_fill = _sample_region_fill(image_path, nodes[0].bbox_px)
        # 像素证据：卡片底色与页面背景色差异大 → 是实色/半透明填充，而非纯描边
        if card_fill:
            cf = np.array([int(card_fill[i:i+2], 16) for i in (1, 3, 5)])
            bg = np.array([int(colors["bg_color"][i:i+2], 16) for i in (1, 3, 5)])
            fill_differs = float(np.linalg.norm(cf - bg)) > 22
    base.style = _build_style(data.get("design", {}), colors, card_fill, fill_differs)

    # 复杂图表区域 → 保留原始像素（VLM 语义门控 + OpenCV 几何定位）。
    # 公式不在此处理：已在节点 body_lines 中按行标记，由布局引擎转 OMML 原生公式。
    vlm_regions = data.get("visual_regions") or []
    keep = detect_keep_regions(image_path, elements, colors["bg_color"],
                               vlm_regions, width, height)
    base.keep_regions = [KeepRegion(bbox_px=r, description="complex_visual") for r in keep]

    # 根治重影：凡与图表保留区重叠的重生成文字，一律剔除（已由保留像素呈现）。
    _strip_overlapping_text(base)
    return base


def _rects_overlap_frac(a, b) -> float:
    """a 被 b 覆盖的面积比例（0~1）。a,b = [x0,y0,x1,y1]。"""
    ix0 = max(a[0], b[0]); iy0 = max(a[1], b[1])
    ix1 = min(a[2], b[2]); iy1 = min(a[3], b[3])
    iw = max(0.0, ix1 - ix0); ih = max(0.0, iy1 - iy0)
    inter = iw * ih
    area_a = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    return inter / area_a


def _bbox_in_any(bbox, regions, thresh=0.5) -> bool:
    """bbox 与任一区域重叠面积 ≥ thresh。regions 为对象列表(.bbox_px)或裸 bbox。"""
    if not bbox or len(bbox) < 4:
        return False
    for r in regions:
        rb = r.bbox_px if hasattr(r, "bbox_px") else r
        if _rects_overlap_frac(bbox, rb) >= thresh:
            return True
    return False


def _strip_overlapping_text(outline: SlideOutline) -> None:
    """剔除与图表保留区重叠的重生成文字，避免与保留像素重影。

    覆盖所有文字载体：title / subtitle / footer / 每个 node。
    """
    blockers = list(outline.keep_regions)
    if not blockers:
        return
    if _bbox_in_any(outline.title_bbox_px, blockers, 0.6):
        outline.title = ""; outline.title_bbox_px = []
    if _bbox_in_any(outline.subtitle_bbox_px, blockers, 0.5):
        outline.subtitle = ""; outline.subtitle_bbox_px = []
    if _bbox_in_any(outline.footer_bbox_px, blockers, 0.5):
        outline.footer_note = ""; outline.footer_bbox_px = []
    kept = [n for n in outline.nodes if not _bbox_in_any(n.bbox_px, blockers, 0.5)]
    dropped = {n.order for n in outline.nodes} - {n.order for n in kept}
    outline.nodes = kept
    outline.connections = [
        c for c in outline.connections
        if c.from_order not in dropped and c.to_order not in dropped
    ]

