# Changelog

## v0.5.0

- Integrated `scipilot-figure-skill` (https://github.com/Haojae/scipilot-figure-skill) as an external dependency for data-driven charts; documented the division of responsibility in SKILL.md: scipilot handles data charts, `generate_diagram.py` handles flowcharts/architecture/block diagrams.
- Added `scripts/generate_diagram.py`: generates flowcharts (`draw_flowchart`), layer/architecture diagrams (`draw_layer_diagram`), and system block diagrams (`draw_block_diagram`) using matplotlib — all three types that scipilot-figure-skill explicitly excludes.
- Added `fit_image_in_box()` in `ppt_utils.py` and as an MCP tool: computes correct (width, height) that fit an image inside a bounding box preserving aspect ratio. Fixes portrait-image overflow that caused height to exceed slide bounds in v2 PPT.
- Added `generate_diagram` MCP tool wrapping all three diagram types.
- Updated SKILL.md Step 5.5: figure strategy table (extract vs. generate), scipilot-figure-skill note, and critical image proportion rule.
- Fixed image proportion issues in the APL paper PPT v2: portrait image `p5_img10.jpeg` (851×1299) was being inserted at 5.8″ wide → 8.9″ tall (overflowing a 7.5″ slide); correct approach constrains by height.

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
