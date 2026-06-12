"""语义大纲 schema：逆推得到的「页面逻辑结构」中间表示。

这是混合重建的核心思想——不逆向像素坐标，而逆向到「大模型生成这页时的结构化大纲」：
布局类型 + 标题 + 节点(内容) + 连接关系。再用原生 PPT 元素重新生成，
得到 100% 可编辑、干净专业的幻灯片，而非像素描摹。

坐标不在这里——VLM 给不准坐标，但能给准结构。坐标由布局引擎按 layout_type 计算。
"""

from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field

from ppt_enhance.builder.font_utils import default_cjk_font


class LayoutType(str, Enum):
    TITLE_COVER = "title_cover"      # 封面
    PROCESS_FLOW = "process_flow"    # 流程图（步骤+箭头，可分支）
    TIMELINE = "timeline"            # 时间线
    COMPARISON = "comparison"        # 对比（并列卡片）
    LIST = "list"                    # 要点列表
    TABLE = "table"                  # 表格
    TREE = "tree"                    # 树/层级
    IMAGE_CONTENT = "image_content"  # 图文混排（含需保留的复杂插画）
    OTHER = "other"


class BodyLine(BaseModel):
    """正文中的一行：可能是普通文字，也可能是公式（需渲染成 OMML）。"""
    text: str = ""
    kind: str = "text"          # text / formula


class OutlineNode(BaseModel):
    """一个逻辑节点（卡片/步骤/要点/表格行）。

    element_ids: 该节点包含的 OCR 文字元素 id —— 用它们的外接框反推节点真实坐标，
    使重建版面与原图一致（而非凭空均分摆位）。
    """
    role: str = "node"
    heading: str = ""           # 小标题
    subtext: str = ""           # 次要文字（作者/来源/副标题）
    body: str = ""              # 说明正文（纯文本拼接，向后兼容）
    body_lines: list["BodyLine"] = Field(default_factory=list)  # 分行正文，区分公式/文字
    order: int = 0
    branch: int = 0             # 分支序号（用于分叉流程；0 为主干）
    element_ids: list[str] = Field(default_factory=list)
    bbox_px: list[float] = Field(default_factory=list)  # 真实像素外接框 [x0,y0,x1,y1]
    children: list["OutlineNode"] = Field(default_factory=list)


class Connection(BaseModel):
    """节点间连接（流程箭头）。"""
    from_order: int
    to_order: int
    style: str = "arrow"


class StyleSpec(BaseModel):
    """自适应样式规格——从原图逆推出的「设计语言」，而非固定模板。

    换一份 PPT，这份规格随之改变，美感自适应原图设计。
    """
    # 卡片
    card_style: str = "translucent"   # solid / translucent / outline
    card_fill: str = "#1C3346"        # 卡片填充色（从图采样）
    card_border: str = "#27A68B"      # 卡片描边色
    card_border_width: float = 1.5
    card_radius: float = 0.12         # 圆角比例（0=方角,0.5=胶囊）
    card_shadow: bool = True
    card_opacity: float = 0.85        # 半透明卡片的不透明度
    # 标题色块（部分设计有顶部强调条）
    header_band: bool = False
    header_band_color: str = "#27A68B"
    # 箭头
    arrow_style: str = "solid"        # solid / thick / gradient
    arrow_width: float = 2.5
    arrow_color: str = "#27A68B"
    # 文字层级
    title_size: float = 28.0
    heading_size: float = 14.0
    body_size: float = 10.5
    # 字体
    font_name: str = Field(default_factory=default_cjk_font)


class KeepRegion(BaseModel):
    """需保留原始像素的复杂插画区域（裁图作为图片对象）。"""
    bbox_px: list[float]        # 原图像素坐标 [x0,y0,x1,y1]
    description: str = ""
    image_path: str | None = None


class FormulaImage(BaseModel):
    """OCR 读成 LaTeX 的公式行，渲染成的公式图（按原 bbox 贴回）。"""
    bbox_px: list[float]        # 原图像素坐标 [x0,y0,x1,y1]
    latex: str = ""             # 原始 OCR 文本
    image_path: str = ""        # 渲染出的透明 PNG 路径


class SlideOutline(BaseModel):
    """单页逆推大纲。"""
    page_no: int
    page_width: float = 0.0      # 原图像素宽（px→英寸映射用）
    page_height: float = 0.0
    source_image_path: str = ""  # 原始渲染图路径（裁保留区像素用）
    layout_type: LayoutType = LayoutType.OTHER
    title: str = ""
    title_bbox_px: list[float] = Field(default_factory=list)
    subtitle: str = ""
    subtitle_bbox_px: list[float] = Field(default_factory=list)
    nodes: list[OutlineNode] = Field(default_factory=list)
    connections: list[Connection] = Field(default_factory=list)
    footer_note: str = ""
    footer_bbox_px: list[float] = Field(default_factory=list)
    keep_regions: list[KeepRegion] = Field(default_factory=list)
    # 主题色（从原图采样，保持品牌一致性）
    bg_color: str = "#0E2A3F"
    accent_color: str = "#27A68B"
    text_color: str = "#FFFFFF"
    style: "StyleSpec" = Field(default_factory=lambda: StyleSpec())
    shapes_summary: str = ""


BodyLine.model_rebuild()
OutlineNode.model_rebuild()
