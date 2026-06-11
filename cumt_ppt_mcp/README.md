# cumt_ppt_mcp

`cumt_ppt_mcp` is a local MCP server for China University of Mining and Technology undergraduate thesis defense PowerPoint automation. It lets Codex inspect and safely modify `.pptx` files through structured tools instead of relying only on natural-language editing.

The first version focuses on CUMT-style defense deck polishing:

- inspect PPTX structure and single-slide details
- apply font rules: title SimHei, body SimSun, Latin/numbers Times New Roman
- normalize right-top CUMT logo size and position
- resize slide images without stretching
- generate thesis-style three-line tables
- export slide PNG previews
- run quality checks and write a markdown report

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
