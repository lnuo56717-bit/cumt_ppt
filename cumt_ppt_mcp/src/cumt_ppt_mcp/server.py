from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from .ppt_utils import (
    PptToolError,
    apply_font_rules as apply_font_rules_impl,
    check_ppt_quality as check_ppt_quality_impl,
    export_preview as export_preview_impl,
    inspect_ppt as inspect_ppt_impl,
    inspect_slide as inspect_slide_impl,
    make_three_line_table as make_three_line_table_impl,
    normalize_logo as normalize_logo_impl,
    resize_images_keep_ratio as resize_images_keep_ratio_impl,
)


mcp = FastMCP("cumt_ppt_mcp")


def _safe_call(fn, *args, **kwargs) -> dict[str, Any]:
    try:
        return fn(*args, **kwargs)
    except PptToolError as exc:
        return {"ok": False, "error": str(exc)}
    except Exception as exc:  # keep MCP tool failures readable for Codex.
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}


@mcp.tool()
def inspect_ppt(pptx_path: str) -> dict[str, Any]:
    """Inspect a PPTX deck: slide count, titles, shape counts, and page size."""
    return _safe_call(inspect_ppt_impl, pptx_path)


@mcp.tool()
def inspect_slide(pptx_path: str, slide_index: int) -> dict[str, Any]:
    """Inspect one 1-based slide and return text, picture, table, and density details."""
    return _safe_call(inspect_slide_impl, pptx_path, slide_index)


@mcp.tool()
def apply_font_rules(
    pptx_path: str,
    output_path: str,
    title_font: str = "黑体",
    body_font: str = "宋体",
    latin_font: str = "Times New Roman",
) -> dict[str, Any]:
    """Apply CUMT defense font rules and save a new PPTX without changing text."""
    return _safe_call(apply_font_rules_impl, pptx_path, output_path, title_font, body_font, latin_font)


@mcp.tool()
def normalize_logo(
    pptx_path: str,
    output_path: str,
    reference_slide_index: int = 3,
    start_slide_index: int = 4,
    skip_slide_indices: list[int] | None = None,
) -> dict[str, Any]:
    """Normalize right-top CUMT logo-like images against a reference slide."""
    return _safe_call(
        normalize_logo_impl,
        pptx_path,
        output_path,
        reference_slide_index,
        start_slide_index,
        skip_slide_indices,
    )


@mcp.tool()
def resize_images_keep_ratio(
    pptx_path: str,
    output_path: str,
    slide_indices: list[int] | None = None,
    mode: str = "fit_content_area",
    max_width_ratio: float = 0.45,
    max_height_ratio: float = 0.58,
) -> dict[str, Any]:
    """Resize selected slide pictures while preserving aspect ratio."""
    return _safe_call(
        resize_images_keep_ratio_impl,
        pptx_path,
        output_path,
        slide_indices,
        mode,
        max_width_ratio,
        max_height_ratio,
    )


@mcp.tool()
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
    """Add a thesis-style three-line table to one slide and save a new PPTX."""
    return _safe_call(make_three_line_table_impl, pptx_path, output_path, slide_index, table_data, left, top, width, height)


@mcp.tool()
def export_preview(pptx_path: str, output_dir: str) -> dict[str, Any]:
    """Export each slide to PNG using PowerPoint COM first, then LibreOffice."""
    return _safe_call(export_preview_impl, pptx_path, output_dir)


@mcp.tool()
def check_ppt_quality(
    pptx_path: str,
    expected_title: str | None = None,
    required_numbers: list[str] | None = None,
    report_path: str | None = None,
) -> dict[str, Any]:
    """Check title/numbers/placeholders/fonts/images/logo consistency and write a markdown report."""
    return _safe_call(check_ppt_quality_impl, pptx_path, expected_title, required_numbers, report_path)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
