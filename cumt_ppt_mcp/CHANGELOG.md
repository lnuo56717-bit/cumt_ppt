# Changelog

## v0.3.1

- Tightened header-logo detection to avoid treating large template content images as CUMT logos.
- Added strict logo candidate constraints for position, size, area ratio, aspect ratio, and body-region exclusion.
- Added template image classification: `header_logo`, `background_decoration`, `content_image`, and `unknown_image`.
- Updated `inspect_template_style` to report image classifications and exclude content figures from reusable style features.
- Updated `apply_template_style` to use only verified `header_logo` placement and the fixed local CUMT logo asset.
- Enhanced layout overlap checks to flag suspected large images misidentified as logos.
- Added regression tests for templates containing a small header logo, a large body figure, and a bottom decoration.

## v0.3.0

- Added fixed CUMT logo asset at `assets/cumt_logo.png`.
- Updated logo normalization to use the fixed local logo asset rather than PDF-extracted or network images.
- Added `check_cumt_logo_asset`.
- Added `add_cumt_logo`.
- Added `inspect_template_style`.
- Added `apply_template_style`.
- Added `check_layout_overlap`.
- Added `auto_fix_layout`.
- Added `extract_pdf_images_safe` for white-background PDF image extraction.
- Added `docs/pdf_image_black_background_investigation.md`.
- Added v0.3.0 tests for logo resource checks, template style inspection, layout overlap detection/fix, and PDF image extraction.

## v0.2.0

- Added `cumt_ppt_mcp`, a Python MCP server for PowerPoint automation.
- Implemented PPTX inspection, slide inspection, font normalization, logo normalization, image resizing, three-line table generation, preview export, and quality checks.
- Added Codex MCP configuration example.
- Added basic pytest tests.
- Verified locally on a real thesis defense PPT without modifying the original file.

## v0.1.0

- Initial `cumt_ppt` skill/prompt workflow for CUMT-style thesis defense PPT generation and polishing.
