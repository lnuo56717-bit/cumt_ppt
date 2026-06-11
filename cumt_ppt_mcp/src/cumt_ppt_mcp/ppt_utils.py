from __future__ import annotations

import math
import os
import re
import shutil
import subprocess
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt


EMU_PER_INCH = 914400
PLACEHOLDER_WORDS = [
    "单击此处",
    "添加标题",
    "添加文本",
    "Key Words",
    "Question",
    "Lorem",
    "Vivamus",
    "Your Text",
    "INSERT",
    "TODO",
]


class PptToolError(ValueError):
    """Clear, user-facing PPT tool error."""


def _path(value: str | os.PathLike[str]) -> Path:
    return Path(value).expanduser().resolve()


def validate_pptx_path(pptx_path: str | os.PathLike[str]) -> Path:
    path = _path(pptx_path)
    if not path.exists():
        raise PptToolError(f"PPTX file does not exist: {path}")
    if not path.is_file():
        raise PptToolError(f"Path is not a file: {path}")
    if path.suffix.lower() != ".pptx":
        raise PptToolError(f"Expected a .pptx file: {path}")
    return path


def validate_output_path(output_path: str | os.PathLike[str], input_path: Path | None = None) -> Path:
    path = _path(output_path)
    if path.suffix.lower() != ".pptx":
        raise PptToolError(f"Output path must end with .pptx: {path}")
    if input_path and path == input_path.resolve():
        raise PptToolError("Output path must be different from input PPTX path.")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def emu_to_inches(value: int | Emu) -> float:
    return round(int(value) / EMU_PER_INCH, 3)


def inches_to_emu(value: float | int) -> Emu:
    return Inches(float(value))


def shape_bounds(shape: Any) -> dict[str, float]:
    return {
        "left": emu_to_inches(shape.left),
        "top": emu_to_inches(shape.top),
        "width": emu_to_inches(shape.width),
        "height": emu_to_inches(shape.height),
    }


def has_text(shape: Any) -> bool:
    return bool(getattr(shape, "has_text_frame", False) and shape.text_frame and shape.text.strip())


def is_picture(shape: Any) -> bool:
    return shape.shape_type == MSO_SHAPE_TYPE.PICTURE


def is_table(shape: Any) -> bool:
    return bool(getattr(shape, "has_table", False))


def iter_text_runs(shape: Any):
    if not has_text(shape):
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            yield paragraph, run


def get_run_font_info(run: Any) -> dict[str, Any]:
    font = run.font
    return {
        "text": run.text,
        "font_name": font.name,
        "size_pt": round(font.size.pt, 1) if font.size else None,
        "bold": font.bold,
        "italic": font.italic,
    }


def shape_font_summary(shape: Any) -> list[dict[str, Any]]:
    fonts = []
    if not has_text(shape):
        return fonts
    for _paragraph, run in iter_text_runs(shape):
        if run.text:
            fonts.append(get_run_font_info(run))
    return fonts


def text_shape_info(shape: Any) -> dict[str, Any]:
    return {
        "text": shape.text.strip(),
        "bounds_in": shape_bounds(shape),
        "fonts": shape_font_summary(shape),
    }


def _font_size_max(shape: Any) -> float:
    sizes = []
    for _paragraph, run in iter_text_runs(shape) or []:
        if run.font.size:
            sizes.append(run.font.size.pt)
    return max(sizes) if sizes else 0.0


def _is_title_shape(shape: Any, slide_height: int) -> bool:
    if not has_text(shape):
        return False
    top_ratio = int(shape.top) / max(slide_height, 1)
    text = shape.text.strip()
    if top_ratio <= 0.16 and len(text) <= 80:
        return True
    return _font_size_max(shape) >= 24 and top_ratio <= 0.28


def detect_title(slide: Any, slide_height: int) -> str:
    candidates = [shape for shape in slide.shapes if _is_title_shape(shape, slide_height)]
    if not candidates:
        candidates = [shape for shape in slide.shapes if has_text(shape)]
    if not candidates:
        return ""
    candidates.sort(key=lambda shape: (shape.top, -_font_size_max(shape), shape.left))
    return candidates[0].text.strip().replace("\n", " ")


def inspect_ppt(pptx_path: str) -> dict[str, Any]:
    path = validate_pptx_path(pptx_path)
    prs = Presentation(str(path))
    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        slides.append(
            {
                "slide_index": idx,
                "title": detect_title(slide, prs.slide_height),
                "text_box_count": sum(1 for shape in slide.shapes if has_text(shape)),
                "image_count": sum(1 for shape in slide.shapes if is_picture(shape)),
                "table_count": sum(1 for shape in slide.shapes if is_table(shape)),
            }
        )
    return {
        "pptx_path": str(path),
        "slide_count": len(prs.slides),
        "page_size": {
            "width_in": emu_to_inches(prs.slide_width),
            "height_in": emu_to_inches(prs.slide_height),
            "width_emu": int(prs.slide_width),
            "height_emu": int(prs.slide_height),
        },
        "slides": slides,
    }


def _dense_hints(slide: Any, prs: Presentation) -> list[str]:
    hints: list[str] = []
    text_shapes = [shape for shape in slide.shapes if has_text(shape)]
    total_chars = sum(len(shape.text.strip()) for shape in text_shapes)
    pictures = [shape for shape in slide.shapes if is_picture(shape)]
    small_fonts = []
    out_of_bounds = []
    for shape in slide.shapes:
        if shape.left < 0 or shape.top < 0 or shape.left + shape.width > prs.slide_width or shape.top + shape.height > prs.slide_height:
            out_of_bounds.append(shape.name)
        for _paragraph, run in iter_text_runs(shape) or []:
            if run.font.size and run.font.size.pt < 10:
                small_fonts.append({"shape": shape.name, "text": run.text[:30], "size_pt": round(run.font.size.pt, 1)})
    if len(text_shapes) >= 12:
        hints.append(f"High number of text boxes: {len(text_shapes)}")
    if total_chars >= 900:
        hints.append(f"Possible dense slide text: {total_chars} characters")
    if len(pictures) >= 6:
        hints.append(f"Many pictures on one slide: {len(pictures)}")
    if small_fonts:
        hints.append(f"Small font risk: {len(small_fonts)} runs under 10 pt")
    if out_of_bounds:
        hints.append(f"Shapes may be outside slide bounds: {', '.join(out_of_bounds[:5])}")
    return hints


def inspect_slide(pptx_path: str, slide_index: int) -> dict[str, Any]:
    path = validate_pptx_path(pptx_path)
    prs = Presentation(str(path))
    if slide_index < 1 or slide_index > len(prs.slides):
        raise PptToolError(f"slide_index must be between 1 and {len(prs.slides)}.")
    slide = prs.slides[slide_index - 1]
    text_boxes = []
    pictures = []
    tables = []
    for shape in slide.shapes:
        if has_text(shape):
            text_boxes.append(text_shape_info(shape))
        if is_picture(shape):
            pictures.append({"name": shape.name, "bounds_in": shape_bounds(shape)})
        if is_table(shape):
            tables.append(
                {
                    "name": shape.name,
                    "bounds_in": shape_bounds(shape),
                    "rows": len(shape.table.rows),
                    "columns": len(shape.table.columns),
                }
            )
    return {
        "pptx_path": str(path),
        "slide_index": slide_index,
        "title": detect_title(slide, prs.slide_height),
        "text_boxes": text_boxes,
        "pictures": pictures,
        "tables": tables,
        "hints": _dense_hints(slide, prs),
    }


def _ensure_rfonts(run: Any, latin_font: str, east_asia_font: str) -> None:
    r_pr = run._r.get_or_add_rPr()
    for tag, face in (("latin", latin_font), ("ea", east_asia_font), ("cs", latin_font)):
        element = r_pr.find(qn(f"a:{tag}"))
        if element is None:
            element = OxmlElement(f"a:{tag}")
            r_pr.append(element)
        element.set("typeface", face)
    run.font.name = latin_font


def apply_font_rules(
    pptx_path: str,
    output_path: str,
    title_font: str = "黑体",
    body_font: str = "宋体",
    latin_font: str = "Times New Roman",
) -> dict[str, Any]:
    input_path = validate_pptx_path(pptx_path)
    out_path = validate_output_path(output_path, input_path)
    prs = Presentation(str(input_path))
    changed_runs = 0
    title_shapes = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not has_text(shape):
                continue
            is_title = _is_title_shape(shape, prs.slide_height)
            if is_title:
                title_shapes += 1
            east_font = title_font if is_title else body_font
            for _paragraph, run in iter_text_runs(shape) or []:
                _ensure_rfonts(run, latin_font, east_font)
                if is_title:
                    run.font.bold = True
                changed_runs += 1
    prs.save(str(out_path))
    return {
        "ok": True,
        "input_path": str(input_path),
        "output_path": str(out_path),
        "changed_runs": changed_runs,
        "title_shapes": title_shapes,
        "title_font": title_font,
        "body_font": body_font,
        "latin_font": latin_font,
    }


def _right_top_pictures(slide: Any, prs: Presentation) -> list[Any]:
    width = int(prs.slide_width)
    height = int(prs.slide_height)
    pics = []
    for shape in slide.shapes:
        if not is_picture(shape):
            continue
        right = int(shape.left + shape.width)
        if shape.left >= width * 0.58 and shape.top <= height * 0.22:
            pics.append(shape)
        elif right >= width * 0.74 and shape.top <= height * 0.25:
            pics.append(shape)
    pics.sort(key=lambda s: (s.top, -s.left))
    return pics


def normalize_logo(
    pptx_path: str,
    output_path: str,
    reference_slide_index: int = 3,
    start_slide_index: int = 4,
    skip_slide_indices: list[int] | None = None,
) -> dict[str, Any]:
    input_path = validate_pptx_path(pptx_path)
    out_path = validate_output_path(output_path, input_path)
    prs = Presentation(str(input_path))
    if reference_slide_index < 1 or reference_slide_index > len(prs.slides):
        raise PptToolError(f"reference_slide_index must be between 1 and {len(prs.slides)}.")
    ref_slide = prs.slides[reference_slide_index - 1]
    ref_pics = _right_top_pictures(ref_slide, prs)
    if not ref_pics:
        raise PptToolError(f"No right-top logo-like image found on reference slide {reference_slide_index}.")
    ref = ref_pics[0]
    skip = set(skip_slide_indices or [])
    modified: list[int] = []
    not_found: list[int] = []
    for idx in range(max(start_slide_index, 1), len(prs.slides) + 1):
        if idx == reference_slide_index or idx in skip:
            continue
        slide = prs.slides[idx - 1]
        pics = _right_top_pictures(slide, prs)
        if not pics:
            not_found.append(idx)
            continue
        logo = pics[0]
        logo.left = ref.left
        logo.top = ref.top
        logo.width = ref.width
        logo.height = ref.height
        modified.append(idx)
    prs.save(str(out_path))
    return {
        "ok": True,
        "input_path": str(input_path),
        "output_path": str(out_path),
        "reference_slide_index": reference_slide_index,
        "reference_logo": shape_bounds(ref),
        "modified_slide_indices": modified,
        "not_found_slide_indices": not_found,
    }


def _selected_slide_indices(total: int, slide_indices: list[int] | None) -> list[int]:
    if not slide_indices:
        return list(range(1, total + 1))
    bad = [idx for idx in slide_indices if idx < 1 or idx > total]
    if bad:
        raise PptToolError(f"slide_indices out of range 1..{total}: {bad}")
    return slide_indices


def _scale_shape_keep_ratio(shape: Any, max_w: int, max_h: int) -> bool:
    ratio = min(max_w / shape.width, max_h / shape.height, 1.0)
    if ratio >= 0.999:
        return False
    shape.width = int(shape.width * ratio)
    shape.height = int(shape.height * ratio)
    return True


def resize_images_keep_ratio(
    pptx_path: str,
    output_path: str,
    slide_indices: list[int] | None = None,
    mode: str = "fit_content_area",
    max_width_ratio: float = 0.45,
    max_height_ratio: float = 0.58,
) -> dict[str, Any]:
    input_path = validate_pptx_path(pptx_path)
    out_path = validate_output_path(output_path, input_path)
    if mode not in {"center", "fit_content_area", "scale"}:
        raise PptToolError("mode must be one of: center, fit_content_area, scale")
    if not 0.05 <= max_width_ratio <= 1.0 or not 0.05 <= max_height_ratio <= 1.0:
        raise PptToolError("max_width_ratio and max_height_ratio must be between 0.05 and 1.0")
    prs = Presentation(str(input_path))
    selected = _selected_slide_indices(len(prs.slides), slide_indices)
    modified: list[dict[str, Any]] = []
    content_top = int(prs.slide_height * 0.16)
    for idx in selected:
        slide = prs.slides[idx - 1]
        pictures = [shape for shape in slide.shapes if is_picture(shape)]
        for pic_shape in pictures:
            old = shape_bounds(pic_shape)
            max_w = int(prs.slide_width * max_width_ratio)
            max_h = int(prs.slide_height * max_height_ratio)
            changed = _scale_shape_keep_ratio(pic_shape, max_w, max_h)
            if mode in {"center", "fit_content_area"}:
                pic_shape.left = int((prs.slide_width - pic_shape.width) / 2)
                if mode == "fit_content_area":
                    pic_shape.top = int(content_top + (prs.slide_height - content_top - pic_shape.height) / 2)
                else:
                    pic_shape.top = int((prs.slide_height - pic_shape.height) / 2)
                changed = True
            elif mode == "scale" and pic_shape.top < content_top:
                pic_shape.top = content_top
                changed = True
            if changed:
                modified.append({"slide_index": idx, "name": pic_shape.name, "old_bounds_in": old, "new_bounds_in": shape_bounds(pic_shape)})
    prs.save(str(out_path))
    return {
        "ok": True,
        "input_path": str(input_path),
        "output_path": str(out_path),
        "mode": mode,
        "modified_images": modified,
    }


def _split_font_runs(text: str) -> list[tuple[str, bool]]:
    parts = re.findall(r"[A-Za-z0-9 .,%+\-_/():=<>×^]+|[^\x00-\x7f]+|.", str(text))
    return [(part, bool(re.fullmatch(r"[A-Za-z0-9 .,%+\-_/():=<>×^]+", part))) for part in parts if part]


def _add_text(slide: Any, left: int, top: int, width: int, height: int, text: str, size_pt: float, bold: bool = False, align: Any = PP_ALIGN.CENTER) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = Inches(0.03)
    frame.margin_right = Inches(0.03)
    frame.margin_top = Inches(0.01)
    frame.margin_bottom = Inches(0.01)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    for part, is_latin in _split_font_runs(text):
        run = paragraph.add_run()
        run.text = part
        _ensure_rfonts(run, "Times New Roman", "宋体")
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(30, 30, 30)
        if is_latin:
            run.font.name = "Times New Roman"


def _add_line(slide: Any, left: int, top: int, width: int, color: RGBColor = RGBColor(29, 67, 113), pt: float = 1.5) -> None:
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Pt(pt))
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.color.rgb = color


def make_three_line_table(
    pptx_path: str,
    output_path: str,
    slide_index: int,
    table_data: list[list[Any]],
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    input_path = validate_pptx_path(pptx_path)
    out_path = validate_output_path(output_path, input_path)
    if not table_data or not all(isinstance(row, list) and row for row in table_data):
        raise PptToolError("table_data must be a non-empty 2D array.")
    columns = len(table_data[0])
    if any(len(row) != columns for row in table_data):
        raise PptToolError("All rows in table_data must have the same number of columns.")
    prs = Presentation(str(input_path))
    if slide_index < 1 or slide_index > len(prs.slides):
        raise PptToolError(f"slide_index must be between 1 and {len(prs.slides)}.")
    slide = prs.slides[slide_index - 1]
    left_emu = inches_to_emu(left if left is not None else 1.0)
    top_emu = inches_to_emu(top if top is not None else 1.35)
    width_emu = inches_to_emu(width if width is not None else 11.3)
    height_emu = inches_to_emu(height if height is not None else 4.6)
    rows = len(table_data)
    row_h = int(height_emu / rows)
    col_w = int(width_emu / columns)
    _add_line(slide, left_emu, top_emu, width_emu, pt=1.6)
    _add_line(slide, left_emu, top_emu + row_h, width_emu, pt=1.2)
    _add_line(slide, left_emu, top_emu + height_emu, width_emu, pt=1.6)
    for r_idx, row in enumerate(table_data):
        for c_idx, value in enumerate(row):
            size = 15 if r_idx == 0 else 13
            bold = r_idx == 0
            _add_text(
                slide,
                left_emu + c_idx * col_w,
                top_emu + r_idx * row_h,
                col_w,
                row_h,
                str(value),
                size,
                bold,
                PP_ALIGN.CENTER,
            )
    prs.save(str(out_path))
    return {
        "ok": True,
        "input_path": str(input_path),
        "output_path": str(out_path),
        "slide_index": slide_index,
        "rows": rows,
        "columns": columns,
        "bounds_in": {
            "left": emu_to_inches(left_emu),
            "top": emu_to_inches(top_emu),
            "width": emu_to_inches(width_emu),
            "height": emu_to_inches(height_emu),
        },
    }


def export_preview(pptx_path: str, output_dir: str) -> dict[str, Any]:
    input_path = validate_pptx_path(pptx_path)
    out_dir = _path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []
    def preview_files() -> list[str]:
        paths = {path.resolve() for pattern in ("*.PNG", "*.png") for path in out_dir.glob(pattern)}
        return [str(path) for path in sorted(paths, key=lambda item: item.name.lower())]

    try:
        import win32com.client  # type: ignore

        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        pres = app.Presentations.Open(str(input_path), False, False, False)
        pres.Export(str(out_dir), "PNG", 1600, 900)
        slide_count = pres.Slides.Count
        pres.Close()
        app.Quit()
        files = preview_files()
        return {
            "ok": True,
            "method": "powerpoint_com",
            "pptx_path": str(input_path),
            "output_dir": str(out_dir),
            "slide_count": slide_count,
            "preview_files": files,
        }
    except Exception as exc:  # pragma: no cover - environment-specific
        errors.append(f"PowerPoint COM export failed: {exc}")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "png", "--outdir", str(out_dir), str(input_path)],
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
            if result.returncode == 0:
                files = preview_files()
                return {
                    "ok": True,
                    "method": "libreoffice",
                    "pptx_path": str(input_path),
                    "output_dir": str(out_dir),
                    "preview_files": files,
                    "stdout": result.stdout,
                }
            errors.append(f"LibreOffice returned {result.returncode}: {result.stderr or result.stdout}")
        except Exception as exc:  # pragma: no cover - environment-specific
            errors.append(f"LibreOffice export failed: {exc}")
    else:
        errors.append("LibreOffice/soffice executable was not found.")
    return {
        "ok": False,
        "pptx_path": str(input_path),
        "output_dir": str(out_dir),
        "errors": errors,
        "suggestion": "Open the PPTX in PowerPoint and export slides as PNG, or install LibreOffice/PowerPoint automation support.",
    }


def _picture_aspect_risk(shape: Any) -> dict[str, Any] | None:
    try:
        blob = shape.image.blob
        with Image.open(BytesIO(blob)) as img:
            image_ratio = img.width / img.height
        shape_ratio = shape.width / shape.height
        if abs(image_ratio - shape_ratio) / image_ratio > 0.08:
            return {
                "name": shape.name,
                "bounds_in": shape_bounds(shape),
                "image_aspect": round(image_ratio, 3),
                "shape_aspect": round(shape_ratio, 3),
            }
    except Exception:
        return None
    return None


def _unique_report_path(pptx_path: Path) -> Path:
    base = pptx_path.with_name(pptx_path.stem + "_quality_report.md")
    if not base.exists():
        return base
    for i in range(1, 100):
        candidate = pptx_path.with_name(f"{pptx_path.stem}_quality_report_{i}.md")
        if not candidate.exists():
            return candidate
    raise PptToolError("Could not create a unique report path.")


def check_ppt_quality(
    pptx_path: str,
    expected_title: str | None = None,
    required_numbers: list[str] | None = None,
    report_path: str | None = None,
) -> dict[str, Any]:
    path = validate_pptx_path(pptx_path)
    prs = Presentation(str(path))
    all_text = []
    small_fonts = []
    aspect_risks = []
    dense_slides = []
    logo_candidates = []
    for idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if has_text(shape):
                all_text.append(shape.text)
                for _paragraph, run in iter_text_runs(shape) or []:
                    if run.font.size and run.font.size.pt < 10:
                        small_fonts.append({"slide_index": idx, "text": run.text[:40], "size_pt": round(run.font.size.pt, 1)})
            if is_picture(shape):
                risk = _picture_aspect_risk(shape)
                if risk:
                    risk["slide_index"] = idx
                    aspect_risks.append(risk)
        hints = _dense_hints(slide, prs)
        if hints:
            dense_slides.append({"slide_index": idx, "hints": hints})
        pics = _right_top_pictures(slide, prs)
        if pics:
            logo_candidates.append((idx, pics[0]))
    text = "\n".join(all_text)
    required_numbers = required_numbers or []
    missing_numbers = [num for num in required_numbers if num not in text]
    placeholders = [word for word in PLACEHOLDER_WORDS if word in text]
    expected_title_present = None if expected_title is None else expected_title in text
    logo_inconsistency = []
    if len(logo_candidates) >= 2:
        ref_idx, ref = logo_candidates[0]
        for idx, shape in logo_candidates[1:]:
            if (
                abs(shape.width - ref.width) > EMU_PER_INCH * 0.03
                or abs(shape.height - ref.height) > EMU_PER_INCH * 0.03
                or abs(shape.left - ref.left) > EMU_PER_INCH * 0.04
                or abs(shape.top - ref.top) > EMU_PER_INCH * 0.04
            ):
                logo_inconsistency.append({"slide_index": idx, "reference_slide_index": ref_idx, "bounds_in": shape_bounds(shape)})
    report_file = _path(report_path) if report_path else _unique_report_path(path)
    if report_file.suffix.lower() != ".md":
        raise PptToolError(f"Quality report path must end with .md: {report_file}")
    report_file.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# PPT Quality Report",
        "",
        f"- PPTX: `{path}`",
        f"- Slides: {len(prs.slides)}",
        f"- Expected title present: {expected_title_present}",
        f"- Missing required numbers: {missing_numbers}",
        f"- Placeholder words found: {placeholders}",
        f"- Small font runs under 10 pt: {len(small_fonts)}",
        f"- Picture aspect risks: {len(aspect_risks)}",
        f"- Logo consistency risks: {len(logo_inconsistency)}",
        "",
        "## Dense Slide Hints",
    ]
    for item in dense_slides:
        lines.append(f"- Slide {item['slide_index']}: {'; '.join(item['hints'])}")
    lines.extend(["", "## Small Fonts"])
    for item in small_fonts[:50]:
        lines.append(f"- Slide {item['slide_index']}: {item['size_pt']} pt, `{item['text']}`")
    lines.extend(["", "## Picture Aspect Risks"])
    for item in aspect_risks[:50]:
        lines.append(f"- Slide {item['slide_index']}: {item['name']} aspect {item['shape_aspect']} vs original {item['image_aspect']}")
    lines.extend(["", "## Logo Risks"])
    for item in logo_inconsistency[:50]:
        lines.append(f"- Slide {item['slide_index']}: differs from slide {item['reference_slide_index']}")
    report_file.write_text("\n".join(lines), encoding="utf-8")
    return {
        "ok": True,
        "pptx_path": str(path),
        "expected_title_present": expected_title_present,
        "missing_required_numbers": missing_numbers,
        "placeholder_words_found": placeholders,
        "small_font_count": len(small_fonts),
        "picture_aspect_risk_count": len(aspect_risks),
        "logo_inconsistency_count": len(logo_inconsistency),
        "dense_slide_hint_count": len(dense_slides),
        "report_path": str(report_file),
    }
