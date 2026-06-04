---
name: cumt_ppt
description: Create and polish China University of Mining and Technology undergraduate thesis defense PowerPoint presentations from thesis PDFs, figure folders, and optional templates, with CUMT-style academic layout, strict font rules, logo consistency, three-line tables, and visual self-check.
---

# CUMT PPT Skill

Use this skill to create or polish China University of Mining and Technology undergraduate thesis defense PowerPoint decks from a thesis PDF, a figure folder, an optional existing PPT, and an optional style template.

## Core Workflow

1. **Collect inputs**
   - Thesis PDF: required for new decks.
   - Figure folder: recommended.
   - Existing PPT: optional, for polishing only.
   - Reference template PPT: optional, for layout/style only.
   - Output path and desired filename.

2. **Extract thesis content**
   - Read the PDF and extract: exact thesis title, author, advisor, college, major, date, abstract, keywords, table of contents, chapter headings, methods, model definitions, algorithm settings, experiment data, conclusions, limitations, and future work.
   - Treat PDF content as highest priority, then existing PPT, then figure filenames/content.
   - Do not invent missing thesis claims or data.

3. **Scan figures**
   - Enumerate `.png`, `.jpg`, `.jpeg`, `.svg`, and other usable image files.
   - Inspect filename, dimensions, clarity, and visual content.
   - Create:
     - `image_inventory.md`: filename, dimensions, likely meaning, clarity, recommended use, reject reason if any.
     - `image_slide_mapping.md`: slide number/title, selected image path, rationale, and fit/crop notes.

4. **Choose a route**
   - **Generate from PDF**: use `prompts/generate_from_pdf.md`.
   - **Polish existing PPT**: use `prompts/polish_existing_ppt.md`; do not rebuild confirmed decks.
   - **Apply optional template**: borrow only layout, font hierarchy, colors, title bars, section rhythm, and table style. Never copy template thesis content, images, data, figures, or unrelated diagrams.
   - **Final cleanup**: use `prompts/final_cleanup.md`.

5. **Build or edit the deck**
   - Recommended page count: 14-20 slides.
   - Use the structure in `references/slide_structure.md` unless the user specifies otherwise.
   - Apply the style rules in `references/style_rules.md`.
   - Apply the font rules in `references/font_rules.md`.
   - Use editable native shapes/tables/charts when image clarity is poor.
   - Use real images only when they match the slide topic and are readable.

6. **Special layout requirements**
   - Keep a unified top title bar and blue-white academic style.
   - Place the China University of Mining and Technology logo consistently at the upper right, using a user-provided or existing in-deck logo. Do not fetch or bundle a new logo.
   - Convert experiment-setting pages, such as PPO parameter pages, into three-line tables when appropriate.
   - For reward ablation, multi-scale, and multi-panel experiment pages, prefer: upper centered charts + lower concise conclusions.
   - Experiment result pages should be evidence-led: conclusion first, then chart/table support.

7. **Validate**
   - Run `scripts/check_ppt_fonts.py` to detect common font violations.
   - Run `scripts/normalize_cumt_logo.py` when logo consistency is required.
   - Run `scripts/export_preview.py` to export PNG previews.
   - Use Computer Use to open PowerPoint and visually inspect every page when the user asks for final visual QA.
   - Produce a visual report with: slide number, check result, modification made, and remaining manual issues.

## Hard Rules

- Never change the thesis title unless the user explicitly asks.
- Never fabricate thesis content, formulas, experiment data, or conclusions.
- Never change experiment data.
- Never place images on unrelated slides.
- Never copy non-user thesis content, images, data, or professional diagrams from a reference template.
- Never use unreadable, blurry, distorted, or topic-mismatched images.
- Never fill slides with long paragraphs; use concise bullets, tables, diagrams, and charts.
- Never rebuild an already accepted PPT unless the user explicitly requests a rebuild.
- Never package user thesis PDFs, formal PPTs, classmate templates, school logo assets, or copyrighted/private materials into this skill.

## Bundled Resources

- `references/font_rules.md`: CUMT defense font hierarchy.
- `references/style_rules.md`: visual style and layout rules.
- `references/slide_structure.md`: recommended slide spine.
- `prompts/generate_from_pdf.md`: prompt for creating a deck from PDF and figures.
- `prompts/polish_existing_ppt.md`: prompt for small-scope polishing.
- `prompts/visual_check.md`: prompt/checklist for PowerPoint visual QA.
- `prompts/final_cleanup.md`: final cleanup checklist.
- `scripts/check_ppt_fonts.py`: report likely font rule violations.
- `scripts/normalize_cumt_logo.py`: normalize upper-right CUMT logo size/position.
- `scripts/export_preview.py`: export PPT pages as PNG previews.
- `scripts/make_three_line_table.py`: helper for drawing editable three-line tables with python-pptx.
