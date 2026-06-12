# PPT Enhance

> 基于多模态解析与多智能体协同的高保真 PDF-to-PPTX 重建与智能纠错系统

将 NotebookLM 等 AI 工具导出的 PDF 演示文稿，转换为**可编辑、已纠错、视觉保真**的 PowerPoint 文件。

## 核心特性

- **多模态解析** — Docling（矢量文本）/ Qwen-OCR（纯图片 PDF）/ MinerU JSON 三条解析通路，提取文本、bbox、图片
- **多智能体纠错** — Contributor 提议 + Reviewer 审查，按修正类别（OCR 纠错 / 去 AI 腔润色）施加差异化阈值，防止过度纠正
- **坐标锚定重建** — python-pptx 按原坐标注入，只改字不改框
- **公式可编辑** — OCR 读成 LaTeX 的公式经 pandoc 转 PowerPoint 原生 OMML，可编辑且排版正确
- **量化评测** — SSIM / PSNR / CER / 可编辑性 + 往返版面 IoU（避免自证的独立保真度测量）

## 快速开始

### 系统依赖（pip 装不上，需单独安装）

两个可选的系统二进制；缺失时对应功能优雅降级（不报错）：

| 工具 | 用途 | 缺失时 | 安装 |
|------|------|--------|------|
| **LibreOffice** | 质量评测（PPTX→PDF 往返渲染算 SSIM/IoU） | 回退占位图，标 `visual_reliable=False` | macOS: `brew install --cask libreoffice`；Ubuntu: `apt install libreoffice` |
| **pandoc** | 公式 LaTeX→OMML 原生可编辑公式 | 公式回退为纯文本 | macOS: `brew install pandoc`；Ubuntu: `apt install pandoc` |

Linux 另需中文字体（否则 outline 版面中文回退替代字体）：`apt install fonts-noto-cjk`。

### 安装与运行

```bash
# 安装
cd PPT_enhance
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"

# 配置 API
cp .env.example .env   # 填入 DashScope/OpenAI 兼容 key

# 生成测试 PDF
python scripts/create_sample_pdf.py

# 命令行转换
ppt-enhance data/samples/sample_slides.pdf

# 纯图片 PDF（如 NotebookLM 导出）走 OCR 通路
ppt-enhance input.pdf --parser qwen-ocr

# 语义大纲逆推模式：用原生 PPT 元素重画，得到干净、100% 可编辑的版面（需 qwen3-vl API key）
ppt-enhance input.pdf --parser qwen-ocr --mode outline

# Web 界面
streamlit run ppt_enhance/ui/app.py
```

> **关于 `--mode outline`（干净可编辑版面）的前提**：该模式靠 VLM（qwen3-vl）逐页逆推语义大纲，
> **必须配置 API key**；未配置时每页得到空白版面（会打印提示，不报错）。纯图片 PDF（NotebookLM
> 等导出）务必同时加 `--parser qwen-ocr`，否则 docling 提取不到文字、版面为空。
> 并发数可用环境变量 `VLM_CONCURRENCY`（默认 6）调整：key 并发额度高时调大提速，过高会限流。

### CLI 参数

| 参数 | 说明 |
|------|------|
| `--parser {docling,qwen-ocr}` | 解析器：矢量文本用 docling，纯图片 PDF 用 qwen-ocr |
| `--mode {anchor,outline}` | 重建路线：`anchor`（坐标锚定，默认，无需 API）/ `outline`（语义大纲逆推，干净可编辑版面，需 qwen3-vl API key） |
| `--mineru-json PATH` | 复用 MinerU JSON（兼容 NotebookLM2PPT 格式） |
| `--no-correction` | 跳过智能纠错 |
| `--no-eval` | 跳过质量评测 |
| `--dpi N` | 渲染 DPI（默认 150） |
| `--no-background` | 关闭整页背景图模式 |
| `--ground-truth PATH` | 提供 ground truth 文本以计算 CER |

## 项目结构

```
ppt_enhance/
├── schemas/      # SlideIR 中间表示 + SlideOutline 语义大纲
├── parser/       # Docling / Qwen-OCR / MinerU / PDF 渲染 / 大纲逆推
├── agents/       # Contributor + Reviewer 多智能体纠错
├── builder/      # python-pptx 重建 + 布局引擎 + 公式 OMML 渲染
├── eval/         # SSIM / PSNR / CER / 往返版面 IoU / 检测框可视化
├── pipeline/     # 主流水线编排
└── ui/           # Streamlit 界面
```

## 当前进展

端到端流水线已跑通：`PDF → 解析 → 多智能体纠错 → 重建 PPTX → 量化评测`，6 个 pytest 全过。

**已实现**
- ✅ 三条解析通路：Docling（矢量）、Qwen-OCR（纯图片 PDF，绝对像素坐标）、MinerU JSON 导入
- ✅ 坐标锚定重建（`pptx_builder`）：按源坐标注入文本，只改字不改框，已接入主流水线
- ✅ 语义大纲逆推重建（`outline_extractor` + `layout_engine`）：不逆向像素而逆向到「生成这页时的语义大纲」，再用原生 PPT 元素重画，得到更干净、100% 可编辑的版面。已接入主流水线，经 `--mode outline` 与坐标锚定路线统一切换（需 qwen3-vl API key；无 key 时回退空白版面）
- ✅ 双 Agent 受控纠错：OCR 纠错走绝对字符差阈值、去 AI 腔润色走比例阈值；跨全文档术语表上下文；无 API key 时规则兜底
- ✅ 公式处理：LaTeX → pandoc → OMML 原生可编辑公式
- ✅ 评测体系：SSIM / PSNR / 可编辑性 / CER，外加往返版面 IoU（PPTX 经 LibreOffice 渲染回 PDF 再独立重提坐标，避免循环论证）
- ✅ 渲染可靠性标记：无 LibreOffice 时回退占位图并标 `visual_reliable=False`，绝不伪造数值
- ✅ Streamlit Web 界面：上传 PDF、配置选项、下载 PPTX/JSON、查看指标与纠错记录
- ✅ 检测框可视化：把 OCR 定位框画回原图，人工核对坐标对齐

**部分实现 / 实验中**
- 🟡 跨平台中文字体：重建时按运行平台自动选择中文字体（macOS PingFang SC / Windows 微软雅黑 / Linux Noto Sans CJK SC），消除「在哪生成、换台机器预览就错位」。Linux 需自备 Noto Sans CJK，否则 LibreOffice 回退等价字体。

## 未实现 / 路线图

按价值优先级排列：

1. **演讲稿生成（规划中，下一步重点）** — 结合用户上传的补充资料 + 各页 PPT 内容，为每页撰写演讲稿，写入 PPTX 的 Speaker Notes。当前代码中无任何 notes 相关实现。详见下方「演讲稿生成」规划。
2. **UI 界面增强（规划中）** — 在现有 Streamlit 基础上扩展：资料上传、逐页预览、演讲稿编辑与导出。详见下方「UI 规划」。
3. **Layout Validator（美学验证器）** — 文本溢出 / 元素碰撞检测，对标 AeSlides。
4. **编辑动作指令范式** — 让 Agent 输出结构化编辑动作而非直接改文本，对标 PPTAgent。
5. **CER 消融实验** — 量化纠错各环节对最终准确率的贡献，用于课程报告。

## 演讲稿生成（规划中）

目标：用户上传补充资料（背景文档、要点提纲、口径要求等），系统结合每页 PPT 的结构化内容，为每页生成演讲稿并写入 Speaker Notes，导出可直接放映。

设计要点（待实现）：
- **输入**：已重建的 SlideIR（含每页文本/标题/逻辑结构） + 用户上传的补充资料（txt/md/pdf/docx）。
- **检索增强**：补充资料切块后按页内容做相关性召回，避免把全文塞进每页 prompt。
- **生成 Agent**：逐页生成讲稿，控制时长口径（如每页 60–90 秒）、衔接上一页、保留专名与数字。
- **落地**：写入 `python-pptx` 的 `slide.notes_slide.notes_text_frame`，并在 Streamlit 中支持逐页查看/编辑/重生成。
- **可选**：整篇讲稿导出为 Markdown / Word；估算总时长。

## UI 规划

在现有 Streamlit 界面（上传 / 配置 / 下载 / 指标 / 纠错记录）基础上扩展：
- 资料上传区（演讲稿用补充材料）。
- 逐页预览：原图 vs 重建效果对照。
- 演讲稿面板：每页讲稿展示、手动编辑、单页重生成、整篇导出。
- 参数预设：针对不同场景（学术 / 商业 / 教学）的纠错与讲稿风格预设。

## 与 NotebookLM2PPT 对比

| 维度 | NotebookLM2PPT | PPT Enhance |
|------|----------------|-------------|
| 平台 | 仅 Windows | 跨平台 |
| 解析 | 微软电脑管家黑盒 OCR | Docling / Qwen-OCR / MinerU 三通路 |
| 纠错 | 无 | 多智能体受控纠错 |
| 公式 | 无 | LaTeX → OMML 可编辑公式 |
| 评测 | 人工目视 | SSIM/PSNR/CER + 往返版面 IoU |
| 演讲稿 | 无 | 规划中（结合用户资料生成 Speaker Notes） |

## 课程项目说明

本项目面向「文本分析与大语言模型」课程期末大作业，强调：
1. 大模型系统设计（多智能体协作）
2. 业务工具开发（PDF→PPT 办公场景）
3. 量化评估（SSIM 视觉保真 + CER 文字准确率 + 往返版面 IoU）
