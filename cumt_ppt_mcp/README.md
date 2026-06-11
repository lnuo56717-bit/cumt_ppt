# cumt_ppt_mcp

`cumt_ppt_mcp` is a local MCP server for China University of Mining and Technology undergraduate thesis defense PowerPoint automation. It lets Codex inspect and safely modify `.pptx` files through structured tools instead of relying only on natural-language editing.

## v0.3.1 Highlights

- Tightens template image classification so large thesis figures, model diagrams, screenshots, and background pictures are not misidentified as right-top CUMT logos.
- Adds strict header-logo candidate rules for top-right placement, maximum width/height, area ratio, aspect ratio, and body-region exclusion.
- `inspect_template_style` now reports picture classifications: `header_logo`, `background_decoration`, `content_image`, and `unknown_image`.
- `apply_template_style` only uses verified `header_logo` placement and the fixed local CUMT logo asset; it does not copy template content images.
- `check_layout_overlap` flags suspected large images misidentified as logos and recommends replacing them with the fixed logo asset.

## v0.3.0 Highlights

- Adds a fixed reusable CUMT logo asset at `assets/cumt_logo.png`, extracted from the final approved defense PPT cover and composited onto a white background to avoid black/transparent rendering problems.
- Adds template-style inspection and conservative template-style application.
- Adds layout overlap detection and conservative auto-fix tools.
- Adds safe PDF image extraction with white-background alpha handling.
- Adds an investigation note at `docs/pdf_image_black_background_investigation.md`.

The first version focuses on CUMT-style defense deck polishing:

- inspect PPTX structure and single-slide details
- apply font rules: title SimHei, body SimSun, Latin/numbers Times New Roman
- normalize right-top CUMT logo size and position
- resize slide images without stretching
- generate thesis-style three-line tables
- export slide PNG previews
- run quality checks and write a markdown report
- inspect template style
- apply template style conservatively
- check and auto-fix layout overlap
- safely extract PDF images without black backgrounds

No thesis PDF, final PPT, student template, logo image, or private material is included in this project. All file paths are supplied by tool parameters.

## Install

```powershell
cd C:\Users\lenovo\Documents\pptskill\cumt_ppt_mcp
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -U pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install -e .
```

For a quick local check without creating a virtual environment:

```powershell
cd C:\Users\lenovo\Documents\pptskill\cumt_ppt_mcp
python -m pip install -r requirements.txt
python -m pip install -e .
```

## Run the MCP Server

```powershell
cd C:\Users\lenovo\Documents\pptskill\cumt_ppt_mcp
python -m cumt_ppt_mcp.server
```

The server uses STDIO transport by default through FastMCP, which is the expected mode for a local Codex MCP server.

## Configure in Codex

Add a local MCP server entry to Codex `config.toml`.

Example:

```toml
[mcp_servers.cumt_ppt_mcp]
command = "python"
args = ["-m", "cumt_ppt_mcp.server"]
cwd = "C:\\Users\\lenovo\\Documents\\pptskill\\cumt_ppt_mcp"
```

If using a virtual environment, point `command` to the venv Python:

```toml
[mcp_servers.cumt_ppt_mcp]
command = "C:\\Users\\lenovo\\Documents\\pptskill\\cumt_ppt_mcp\\.venv\\Scripts\\python.exe"
args = ["-m", "cumt_ppt_mcp.server"]
cwd = "C:\\Users\\lenovo\\Documents\\pptskill\\cumt_ppt_mcp"
```

See `examples/codex_config_example.toml`.

## Available Tools

### `inspect_ppt`

Input:

- `pptx_path`

Returns slide count, page size, each slide title, text box count, image count, and table count.

### `inspect_slide`

Input:

- `pptx_path`
- `slide_index` 1-based

Returns text boxes, positions, font information, images, tables, and density warnings.

### `apply_font_rules`

Input:

- `pptx_path`
- `output_path`
- `title_font`, default `黑体`
- `body_font`, default `宋体`
- `latin_font`, default `Times New Roman`

Always writes to `output_path`; it never overwrites the input file.

### `normalize_logo`

Input:

- `pptx_path`
- `output_path`
- `reference_slide_index`, default `3`
- `start_slide_index`, default `4`
- `skip_slide_indices`, optional

Uses the right-top image on the reference slide as the logo standard and aligns later right-top logo-like images to it.

### `resize_images_keep_ratio`

Input:

- `pptx_path`
- `output_path`
- `slide_indices`, optional
- `mode`: `center`, `fit_content_area`, or `scale`
- `max_width_ratio`
- `max_height_ratio`

Keeps each image aspect ratio and avoids stretching.

### `make_three_line_table`

Input:

- `pptx_path`
- `output_path`
- `slide_index`
- `table_data`, 2D array
- `left`, `top`, `width`, `height`, optional, in inches

Adds a thesis-style three-line table to the slide.

### `export_preview`

Input:

- `pptx_path`
- `output_dir`

Tries PowerPoint COM first on Windows, then LibreOffice. If neither works, returns a clear error and suggestion.

### `check_ppt_quality`

Input:

- `pptx_path`
- `expected_title`, optional
- `required_numbers`, optional array
- `report_path`, optional markdown output path

Checks title/numbers/placeholders/small fonts/image aspect risks/logo consistency and writes a markdown report. If `report_path` is omitted, it creates a unique report next to the PPTX.

### `check_cumt_logo_asset`

Input:

- `logo_asset_path`, optional

Checks the fixed CUMT logo asset for black background, transparency, and size risks.

### `add_cumt_logo`

Input:

- `pptx_path`
- `output_path`
- `slide_indices`, optional
- `logo_asset_path`, optional, defaults to `assets/cumt_logo.png`
- `left`, `top`, `width`, optional, in inches
- `skip_slide_indices`, optional

Adds the fixed local CUMT logo asset to selected slides.

### `inspect_template_style`

Input:

- `template_pptx_path`

Returns page size, common colors, title bar candidates, title/body font style, layout type guesses, verified logo positions, page number positions, emphasis colors, element position parameters, and an `image_classifications` table. Large content figures are classified as `content_image` and are excluded from reusable style features.

### `apply_template_style`

Input:

- `source_pptx_path`
- `template_pptx_path`
- `output_path`

Conservatively applies template title bars, font hierarchy, margins, and fixed CUMT logo placement without changing slide text. It uses only verified `header_logo` placement and never copies large template content pictures into the target deck.

### `check_layout_overlap`

Input:

- `pptx_path`
- `report_path`, optional

Checks text, image, table, and logo objects for overlap, out-of-bounds placement, title obstruction, logo/title conflicts, oversized logos, and suspected large images misidentified as logos.

### `auto_fix_layout`

Input:

- `pptx_path`
- `output_path`
- `report_path`, optional

Conservatively fixes obvious image overlap and out-of-bounds issues while preserving image ratio and text content.

### `extract_pdf_images_safe`

Input:

- `pdf_path`
- `output_dir`
- `background_color`, default `white`

Extracts embedded PDF images and page renders, composites transparency onto a white background, and writes an inventory plus report.

## Safety Rules

- No tool executes arbitrary shell commands.
- No tool deletes user files.
- No tool overwrites the input PPTX.
- All write tools require an explicit `output_path`, except quality reports, which create a new unique report path next to the PPTX.
- Input paths must exist.
- PPT files must use the `.pptx` extension.
- Errors are returned as clear `{"ok": false, "error": "..."}` results.

## Test

```powershell
cd C:\Users\lenovo\Documents\pptskill\cumt_ppt_mcp
python -m pytest
```

The tests create temporary PPTX files and do not use private thesis files.
