# Changelog

## v0.4.0

- Extended SKILL.md to support journal/conference paper presentations alongside thesis defense.
- Added `render_formula_png()` in `ppt_utils.py`: renders LaTeX/mathtext formulas as transparent PNGs using matplotlib mathtext (no LaTeX install required).
- Added `add_breadcrumb_strip()` in `ppt_utils.py`: draws a per-section navigation strip on a slide with an active-section amber highlight.
- Exposed both as `@mcp.tool()` entries in `server.py`.
- Added `scripts/render_formulas.py`: standalone script for batch formula rendering with color and sizing options.
- Added `prompts/generate_from_journal.md`: end-to-end prompt for creating a deck from a journal paper PDF.
- Updated `references/slide_structure.md` with a 12–16 slide journal paper spine and breadcrumb section reference.
- Documented font pitfalls: set East Asian font via `a:ea`/`a:cs` XML elements, not `rf.name`; avoid smart/curly quote characters in Python string literals.
- Verified locally on an APL journal paper (superconducting infrared nanobolometer) with figure extraction, formula rendering, and breadcrumb navigation.

## v0.3.1

- Fixed false detection of large template content images as header logos.
- Added stricter logo candidate rules based on position, size, area ratio, aspect ratio, and content-region exclusion.
- Added image classification in template inspection: `header_logo`, `background_decoration`, `content_image`, and `unknown_image`.
- Updated `apply_template_style` to apply reusable style features without copying large content images from templates.
- Improved overlap checking for oversized logos, title-area obstruction, and suspected logo misclassification.
- Added regression tests for template image classification and logo candidate detection.
- Verified locally with an external thesis PDF and template PPT.
- Kept all real PDFs, PPTs, template files, preview images, and test outputs excluded from version control.

## v0.3.0

- Added fixed CUMT logo asset support for `cumt_ppt_mcp`.
- Added template-style inspection and conservative template-style application.
- Added layout overlap detection and conservative auto-fix tools.
- Added safe PDF image extraction with white-background alpha handling.
- Added PDF black-background investigation notes.

## v0.2.0

- Added `cumt_ppt_mcp`, a Python MCP server for PowerPoint automation.
- Implemented PPTX inspection, slide inspection, font normalization, logo normalization, image resizing, three-line table generation, preview export, and quality checks.
- Added Codex MCP configuration example.
- Added basic pytest tests.
- Verified locally on a real thesis defense PPT without modifying the original file.
- Excluded real thesis files, PPT files, preview images, and test outputs from version control.

## v0.1.0

- Initial `cumt_ppt` skill/prompt workflow for CUMT-style thesis defense PPT generation and polishing.
