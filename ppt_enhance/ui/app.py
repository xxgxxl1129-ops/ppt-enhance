"""Streamlit 交互界面."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

import streamlit as st

from ppt_enhance.pipeline.main_pipeline import run_pipeline


st.set_page_config(
    page_title="PPT Enhance",
    page_icon="📊",
    layout="wide",
)

st.title("📊 PPT Enhance")
st.caption("高保真 PDF-to-PPTX 重建 · 多智能体智能纠错 · 量化评测")

with st.sidebar:
    st.header("⚙️ 设置")
    parser_choice = st.radio(
        "解析器",
        options=["docling", "qwen-ocr"],
        format_func=lambda p: {"docling": "Docling（矢量文本 PDF）",
                               "qwen-ocr": "Qwen-OCR（纯图片 PDF，如 NotebookLM 导出）"}[p],
        help="纯图片 PDF（NotebookLM 等导出）必须用 Qwen-OCR，否则提取不到文字、版面为空。",
    )
    mode = st.radio(
        "重建路线",
        options=["anchor", "outline"],
        format_func=lambda m: {"anchor": "坐标锚定（默认，无需 API）",
                               "outline": "语义大纲（干净可编辑，需 qwen3-vl API）"}[m],
        help="outline 模式用 VLM 逆推每页语义大纲再用原生元素重画，版面更干净；未配置 API key 时会得到空白页。",
    )
    enable_correction = st.toggle("启用智能纠错", value=True)
    enable_eval = st.toggle("启用质量评测", value=True)
    use_background = st.toggle("整页背景模式（高视觉保真）", value=True,
                               help="仅 anchor 模式生效")
    dpi = st.slider("渲染 DPI", 72, 300, 150, step=10)
    mineru_json = st.file_uploader("MinerU JSON（可选）", type=["json"])
    ground_truth = st.text_area("Ground Truth 文本（CER 评测，可选）", height=100)

uploaded = st.file_uploader("上传 PDF 文件", type=["pdf"])

if uploaded and st.button("🚀 开始转换", type="primary"):
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        pdf_path = tmp_path / uploaded.name
        pdf_path.write_bytes(uploaded.read())

        mineru_path = None
        if mineru_json:
            mineru_path = tmp_path / "mineru.json"
            mineru_path.write_bytes(mineru_json.read())

        output_dir = tmp_path / "output"
        progress = st.progress(0, text="正在处理...")

        # 进度回调由 run_pipeline 在主线程触发（VLM 并行在工作线程，
        # 但回调统一回到主线程），故可直接更新 st.progress。
        def on_progress(frac: float, text: str) -> None:
            progress.progress(int(frac * 100), text=text)

        try:
            result = run_pipeline(
                pdf_path=pdf_path,
                output_dir=output_dir,
                mineru_json=mineru_path,
                enable_correction=enable_correction,
                enable_eval=enable_eval,
                dpi=dpi,
                use_background=use_background,
                ground_truth_text=ground_truth or None,
                mode=mode,
                parser=parser_choice,
                progress_cb=on_progress,
            )
            progress.progress(100, text="完成")

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("📥 下载结果")
                pptx_bytes = result.pptx_path.read_bytes()
                st.download_button(
                    "下载 PPTX",
                    data=pptx_bytes,
                    file_name=result.pptx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
                ir_bytes = (output_dir / "slide_ir_corrected.json").read_bytes()
                st.download_button("下载 SlideIR JSON", data=ir_bytes, file_name="slide_ir_corrected.json")

            with col2:
                st.subheader("📈 质量指标")
                if result.eval_report:
                    er = result.eval_report
                    m1, m2, m3 = st.columns(3)
                    m1.metric("SSIM", f"{er.ssim_mean:.4f}")
                    m2.metric("PSNR", f"{er.psnr_mean:.1f} dB")
                    m3.metric("可编辑性", f"{er.editability_ratio:.1%}")
                    if er.cer is not None:
                        st.metric("CER", f"{er.cer:.4f}")
                    for note in er.notes:
                        st.info(note)

            if result.correction_records:
                st.subheader("✏️ 纠错记录")
                accepted = [r for r in result.correction_records if r.accepted]
                rejected = [r for r in result.correction_records if not r.accepted]

                tab_a, tab_r = st.tabs([f"✅ 已通过 ({len(accepted)})", f"❌ 已拒绝 ({len(rejected)})"])
                with tab_a:
                    for r in accepted:
                        st.markdown(f"**{r.element_id}**")
                        st.markdown(f"- 原文: `{r.original}`")
                        st.markdown(f"- 修正: `{r.corrected}`")
                        st.markdown(f"- 原因: {r.reason}")
                with tab_r:
                    for r in rejected:
                        st.markdown(f"**{r.element_id}**: {r.reject_reason}")

            st.subheader("📋 解析概览")
            summary = {
                "pages": len(result.slide_ir.pages),
                "elements": sum(len(p.elements) for p in result.slide_ir.pages),
                "text_elements": len(result.slide_ir.all_text_elements()),
                "protected_terms": result.slide_ir.protected_terms[:10],
            }
            st.json(summary)

        except Exception as e:
            progress.empty()
            st.error(f"转换失败: {e}")
            st.exception(e)

else:
    st.info("请上传 PDF 文件后点击「开始转换」。")
    st.markdown("""
    ### 系统架构
    1. **VLM 解析** — Docling 提取文本 + bbox + 图片
    2. **多智能体纠错** — Contributor 提议 + Reviewer 审查
    3. **坐标锚定重建** — python-pptx 按原坐标注入，只改字不改框
    4. **量化评测** — SSIM / PSNR / CER / 可编辑性
    """)
