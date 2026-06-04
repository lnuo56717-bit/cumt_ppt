# Generate From PDF Prompt

Use this prompt when creating a CUMT undergraduate thesis defense PPT from a thesis PDF and figure folder.

## User Inputs

- Thesis PDF:
- Figure folder:
- Optional reference template:
- Output PPT path:
- Desired page count:
- Student/advisor/college/major/date:

## Instructions

1. Read the thesis PDF first.
2. Extract the exact thesis title and never change it.
3. Extract author, advisor, college, major, abstract, keywords, table of contents, chapter headings, methods, algorithms, experiment settings, data, conclusions, limitations, and future work.
4. Scan the figure folder and build:
   - `image_inventory.md`
   - `image_slide_mapping.md`
5. Create a 14-20 page defense PPT using `references/slide_structure.md`.
6. Use a CUMT blue-white academic style:
   - unified top title bar;
   - CUMT logo at upper right if a user-provided or in-deck logo is available;
   - concise bullets;
   - editable diagrams and three-line tables;
   - upper centered charts + lower conclusions for multi-chart experiment pages.
7. If a reference template is provided, borrow only layout, font hierarchy, colors, and section rhythm. Do not copy its content, images, data, or diagrams.
8. Normalize fonts using `references/font_rules.md`.
9. Export previews and produce a final report.

## Mandatory Checks

- Exact thesis title preserved.
- No fabricated content or data.
- Experiment data unchanged.
- Images match slide topics.
- Unreadable images replaced by native tables/diagrams.
- Template content not copied.
- PPT opens successfully.
