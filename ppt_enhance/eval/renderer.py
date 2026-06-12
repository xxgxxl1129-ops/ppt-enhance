"""将 PPTX 渲染为页面图像（用于 SSIM 评测）."""

from __future__ import annotations

from pathlib import Path

import fitz
from pptx import Presentation


def pptx_to_images(
    pptx_path: str | Path,
    output_dir: str | Path,
    dpi: int = 150,
) -> tuple[list[Path], bool, Path | None]:
    """将 PPTX 每页导出为 PNG。

    返回 (图片路径列表, is_real, rendered_pdf_path)。
    is_real=True 表示由 LibreOffice 真实渲染，SSIM/PSNR 可信；
    is_real=False 表示回退到占位图，视觉指标不可信。
    rendered_pdf_path: LibreOffice 渲染出的 PDF（供往返版面 IoU 独立重解析）；
    无 LibreOffice 时为 None。
    """
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 尝试 libreoffice 转 PDF
    pdf_path = output_dir / f"{pptx_path.stem}.pdf"
    if _try_libreoffice_convert(pptx_path, pdf_path):
        return _pdf_to_pngs(pdf_path, output_dir, dpi), True, pdf_path

    # 回退：用幻灯片背景+元素近似渲染（基于已有 render 图）
    return _fallback_render(pptx_path, output_dir), False, None


def _try_libreoffice_convert(pptx_path: Path, pdf_path: Path) -> bool:
    import shutil
    import subprocess

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(pptx_path)],
            check=True,
            capture_output=True,
            timeout=120,
        )
        return pdf_path.exists()
    except Exception:
        return False


def _pdf_to_pngs(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    from ppt_enhance.parser.pdf_renderer import render_pdf_pages

    pages = render_pdf_pages(pdf_path, output_dir / "pptx_pages", dpi=dpi)
    return [pages[k] for k in sorted(pages)]


def _fallback_render(pptx_path: Path, output_dir: Path) -> list[Path]:
    """无 LibreOffice 时，导出幻灯片数量对应的占位图（基于幻灯片尺寸）."""
    from PIL import Image, ImageDraw, ImageFont

    prs = Presentation(str(pptx_path))
    paths: list[Path] = []
    w = int(prs.slide_width / 914400 * 150)  # EMU to px approx at 150dpi
    h = int(prs.slide_height / 914400 * 150)

    for i, _slide in enumerate(prs.slides):
        img = Image.new("RGB", (max(w, 1280), max(h, 720)), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        # 占位图只画英文标记，无需 CJK 字体；任一平台失败均回退默认字体
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", 24)
        except OSError:
            font = ImageFont.load_default()
        draw.text((40, 40), f"Slide {i + 1} (fallback render)", fill=(100, 100, 100), font=font)
        out = output_dir / f"slide_{i + 1:04d}.png"
        img.save(out)
        paths.append(out)
    return paths
