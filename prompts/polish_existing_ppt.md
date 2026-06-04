# Polish Existing PPT Prompt

Use this prompt when the user says an existing PPT is basically usable and wants small-scope improvements.

## Inputs

- Existing PPT:
- Thesis PDF, optional:
- Figure folder, optional:
- Reference template, optional:
- Output PPT path:

## Instructions

1. Do not rebuild the PPT unless explicitly requested.
2. Do not change page count, page order, thesis title, experiment data, or confirmed image mapping.
3. Identify only format, spacing, logo, font, table, and image-size issues.
4. Apply small changes:
   - normalize title/body/English fonts;
   - align top title bars and page numbers;
   - unify CUMT logo position and size;
   - convert parameter cards into three-line tables;
   - center and enlarge experiment charts if needed;
   - move concise conclusions below multi-chart figures;
   - preserve image ratios.
5. Export PNG previews.
6. Write a per-slide modification report.

## Hard Stop Conditions

Ask before rebuilding the deck, changing core content, changing experiment values, replacing images, or copying any reference-template content.
