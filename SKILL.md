---
name: cumt_ppt
description: Create and polish CUMT-style academic PowerPoint presentations from thesis PDFs or journal papers, with section breadcrumb navigation, embedded PDF figure extraction, formula rendering, blue-white academic style, strict font rules, three-line tables, and visual self-check.
---

# CUMT PPT Skill

Use this skill to create CUMT-style academic PowerPoint decks from a thesis PDF or a journal/conference paper PDF, with optional figure folders and reference templates.

## Supported Modes

| Mode | When to use | Prompt to load |
|------|-------------|----------------|
| Generate from thesis PDF | CUMT undergraduate/graduate thesis defense | `prompts/generate_from_pdf.md` |
| Generate from journal paper | APL / Nature / IEEE journal article presentation | `prompts/generate_from_journal.md` |
| Polish existing PPT | Small-scope improvement of an accepted deck | `prompts/polish_existing_ppt.md` |

## Core Workflow

1. **Collect inputs**
   - Source PDF: required (thesis or journal paper).
   - Figure folder: optional; if absent, extract figures directly from PDF.
   - Existing PPT: optional, for polishing only.
   - Reference template PPT: optional, for layout/style only.
   - Output path and desired filename.

2. **Extract content from PDF**
   - Read the full PDF text and extract:
     - Thesis: title, author, advisor, college, major, date, abstract, keywords, TOC, chapter headings, methods, algorithms, experiment data, conclusions, limitations, future work.
     - Journal paper: title, authors, affiliations, journal/DOI, abstract, introduction motivation, device/method design, key equations, simulation results, experimental setup, measured data, conclusions.
   - Treat PDF text as highest priority. Do not fabricate content, formulas, or data.

3. **Extract figures from PDF**
   - **Embedded images first**: use `extract_pdf_images_safe` MCP tool (or PyMuPDF `page.get_images()`) to pull high-resolution images directly embedded in the PDF. These are typically the paper's actual figures and are higher quality than page renders.
   - **Page renders as fallback**: if embedded images are too small or absent, render full pages at 3× zoom (~216 DPI) using `fitz.Matrix(3, 3)`.
   - Record a figure inventory: page number, image index, pixel dimensions, likely figure label (Fig. 1, Fig. 2 …), and recommended slide target.
   - Skip cover/TOC/reference page images (logos, icons < 200 px).

4. **Render formulas as images**
   - For any mathematical formula mentioned in the paper, render it as a transparent PNG using `scripts/render_formulas.py` (matplotlib mathtext, no LaTeX install required).
   - Key rules for matplotlib mathtext:
     - Use a single `$...$` block per `ax.text` call; avoid adjacent `$...$` pairs.
     - Do not use `\&` (unsupported); use `\text{and}` or plain text instead.
     - `\dfrac`, `\sqrt`, `\times`, `\mathcal`, `\mathrm`, `\Rightarrow` all work.
     - Save with `transparent=True`, `bbox_inches='tight'`, `dpi=160`.
   - Insert formula PNGs into slides using `slide.shapes.add_picture()`.

5. **Choose slide structure**
   - For thesis defense: use `references/slide_structure.md` (14–20 slides).
   - For journal paper: use `references/slide_structure.md` journal section (12–16 slides):
     1. Cover (title, authors, journal, DOI)
     2. Outline
     3. Research background & motivation
     4. Device / method design
     5. Working principle
     6. Simulation analysis
     7. Fabrication & characterization
     8. Experimental setup
     9. Key results (absorption / response / spectra)
     10. Performance metrics table
     11. Conclusions & outlook
     12. Acknowledgements / Thank you

5.5. **Figure strategy — what to generate vs. what to extract**

   Two figure sources coexist; choose per slide type:

   | Slide content | Source | Tool |
   |---|---|---|
   | Actual measurement spectra, SEM/TEM images, R-T curves | Extract from PDF | `extract_pdf_images_safe` MCP / PyMuPDF |
   | Data charts (absorption, NEP, D*, responsivity vs. frequency) | Generate from paper data | `scipilot-figure-skill` (see below) |
   | Process flow, research methodology, decision logic | **Generate** | `scripts/generate_diagram.py` `draw_flowchart()` |
   | System/device layer structure | **Generate** | `scripts/generate_diagram.py` `draw_layer_diagram()` |
   | Measurement setup block diagram | **Generate** | `scripts/generate_diagram.py` `draw_block_diagram()` |

   **scipilot-figure-skill** ([github.com/Haojae/scipilot-figure-skill](https://github.com/Haojae/scipilot-figure-skill)):
   - Use for data-driven charts: line plots, bar charts, scatter, heatmaps, violin plots, error bars.
   - Follow its 8-step workflow: profile data → select chart type → configure style → draw → visual-QA loop → export.
   - Note: it explicitly does **not** generate flowcharts, architecture diagrams, or schematics — use `generate_diagram.py` for those.

   **Image proportion rule (critical)**:
   - Always use `fit_image_in_box(img_path, max_w_in, max_h_in)` from `generate_diagram.py` to get the display width/height before calling `slide.shapes.add_picture()`.
   - Portrait images (height > width) at a fixed wide `w` will overflow the slide — let `fit_image_in_box` scale them down by height instead.
   - Example: a 851×1299 image inserted at 5.8″ wide → 8.9″ tall (overflows 7.5″ slide). Correct: clamp to max_h=5.5″ → 3.6″ wide.

6. **Build the deck (python-pptx)**
   - Slide size: 13.333″ × 7.5″ (widescreen).
   - Apply style from `references/style_rules.md` and fonts from `references/font_rules.md`.
   - **Title bar** (top, 1.0″ tall, deep blue): slide number badge + title text.
   - **Breadcrumb navigation strip** (below title bar, 0.32″ tall): show all outline sections as colored tabs; highlight the current section in amber/gold. Use `add_breadcrumb_strip` MCP tool or implement directly with `add_shape` rectangles.
   - **Content area**: starts at 1.32″ from top.
   - **Font rules (critical)**:
     - In `run.font`, set `rf.name = en_font` for Latin.
     - Then set `a:ea typeface = cn_font` via XML (`run._r.get_or_add_rPr()`) for Chinese characters.
     - Setting `rf.name = cn_font` then `rf.name = en_font` will overwrite the Chinese font — always set East Asian via XML.
   - **Chinese string literals**: avoid smart/curly quotes `"` `"` (U+201C/U+201D) inside Python string literals; they are indistinguishable from ASCII `"` in some editors and cause SyntaxError. Use `[` `]` or `'` `'` instead.
   - Layout per slide type:
     - Left-text / right-image: text 45–52 %, image 45–52 %.
     - Conclusion-first experiment page: key finding box on top, chart/table below.
     - Three-line table for parameter/settings pages.

7. **Special layout requirements**
   - Keep a unified top title bar and blue-white academic style.
   - Breadcrumb must reflect the active outline section on every content slide.
   - Place the CUMT logo consistently at upper right when a logo asset is available.
   - Convert experiment-setting pages into three-line tables when appropriate.
   - For multi-panel figure pages, prefer: upper centered charts + lower concise conclusions.

8. **Validate**
   - Run `scripts/check_ppt_fonts.py` to detect font violations.
   - Run `scripts/normalize_cumt_logo.py` when logo consistency is required.
   - Run `scripts/export_preview.py` to export PNG previews.
   - Run `check_ppt_quality` MCP tool for title, placeholder, and font checks.
   - Produce a final report: slide number, check result, modification made, remaining issues.

## Hard Rules

- Never change the original paper/thesis title unless the user explicitly asks.
- Never fabricate content, formulas, experiment data, or conclusions.
- Never place figures on slides unrelated to those figures.
- Never copy content, images, or data from a reference template.
- Never use unreadable, blurry, or topic-mismatched images.
- Never fill slides with dense paragraphs; use bullets, tables, and charts.
- Never rebuild an accepted deck unless the user explicitly requests it.
- Never package private PDFs, PPTs, thesis files, or copyrighted materials into this skill.

## External Skill Dependencies

- **scipilot-figure-skill** (`https://github.com/Haojae/scipilot-figure-skill`): for publication-quality data charts. Covers matplotlib + seaborn, journal specs, visual QA loop, CJK font fix. Does NOT generate flowcharts or architecture diagrams.

## Bundled Resources

- `references/font_rules.md`: CUMT defense font hierarchy.
- `references/style_rules.md`: visual style and layout rules.
- `references/slide_structure.md`: recommended slide spines (thesis + journal paper).
- `prompts/generate_from_pdf.md`: prompt for creating a deck from a thesis PDF.
- `prompts/generate_from_journal.md`: prompt for creating a deck from a journal paper.
- `prompts/polish_existing_ppt.md`: prompt for small-scope polishing.
- `prompts/visual_check.md`: prompt/checklist for PowerPoint visual QA.
- `prompts/final_cleanup.md`: final cleanup checklist.
- `scripts/check_ppt_fonts.py`: report likely font rule violations.
- `scripts/normalize_cumt_logo.py`: normalize upper-right CUMT logo size/position.
- `scripts/export_preview.py`: export PPT pages as PNG previews.
- `scripts/make_three_line_table.py`: helper for drawing editable three-line tables.
- `scripts/render_formulas.py`: render LaTeX/mathtext formulas as transparent PNG images.
- `scripts/generate_diagram.py`: generate flowcharts, layer diagrams, and block diagrams as PNGs (covers what scipilot-figure-skill explicitly excludes). Run `--demo` to preview three sample diagrams.
