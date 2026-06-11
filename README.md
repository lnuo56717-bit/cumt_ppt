# cumt_ppt

`cumt_ppt` is a Codex Skill for creating and polishing China University of Mining and Technology undergraduate thesis defense PowerPoint presentations.

It helps Codex turn a thesis PDF, a figure folder, an optional existing PPT, and an optional reference template into a CUMT-style academic defense deck with:

- blue-white academic layout;
- strict Chinese/English font rules;
- unified CUMT logo placement;
- three-line experiment tables;
- figure-to-slide matching;
- PowerPoint visual self-check and PNG previews.

## Inputs To Prepare

- Thesis PDF, preferably the final or near-final version.
- Figure folder with charts, layouts, flowcharts, screenshots, and experiment figures.
- Optional existing PPT when you want polishing instead of rebuilding.
- Optional reference template PPT for style only.
- Output filename and page count preference.

Do not include confidential or copyrighted files in this skill folder. Keep user PDFs, formal PPTs, classmate templates, logos, and private datasets outside the skill.

## Recommended Workflow

1. Ask Codex to use `cumt_ppt`.
2. Provide the thesis PDF path, figure folder path, optional existing PPT, optional template, and output path.
3. Codex extracts thesis content and scans figures.
4. Codex creates `image_inventory.md` and `image_slide_mapping.md` in the working project folder.
5. Codex generates or polishes the PPT.
6. Codex normalizes fonts, logo placement, tables, and multi-figure layouts.
7. Codex opens the deck in PowerPoint with Computer Use for visual checking when requested.
8. Codex exports PNG previews and writes a visual report.

## v0.2.0 - cumt_ppt_mcp

This repository now includes `cumt_ppt_mcp`, a Python MCP tool server that lets Codex directly inspect and operate on PPTX files.

Implemented tools:

- `inspect_ppt`
- `inspect_slide`
- `apply_font_rules`
- `normalize_logo`
- `resize_images_keep_ratio`
- `make_three_line_table`
- `export_preview`
- `check_ppt_quality`

The MCP server has been verified locally on a real thesis defense PPT without modifying the original file. Real PPT files, thesis PDFs, preview images, and local test outputs are excluded from version control.

Usage details are in [`cumt_ppt_mcp/README.md`](cumt_ppt_mcp/README.md).

## How To Ask Codex To Use It

Example:

```text
Use the cumt_ppt skill. Create a CUMT undergraduate thesis defense PPT from:
- PDF: C:\path\to\thesis.pdf
- Figures: C:\path\to\figures
- Optional template: C:\path\to\template.pptx
- Output: C:\path\to\defense_v1.pptx
Do not invent data and do not copy template content.
```

For polishing:

```text
Use the cumt_ppt skill. Polish this existing PPT only:
C:\path\to\defense_v1.pptx
Do not rebuild or change content. Normalize fonts, logo, tables, image alignment, and export previews.
```

## Uploading To GitHub

If this skill should be shared through GitHub:

1. Create an empty repository on GitHub.
2. From the parent folder of `cumt_ppt`, run:

```bash
git remote add origin <your-repo-url>
git push -u origin main
```

Only upload this generic skill folder. Do not upload:

- personal thesis PDFs;
- formal defense PPTs;
- classmate PPT templates;
- school logo image files copied from official or private sources;
- copyrighted or confidential figures.

## Privacy And Copyright

This skill contains only reusable rules, prompts, and helper scripts. It intentionally does not include user papers, templates, or image assets. When using reference templates, borrow layout and style only; never copy their thesis content, diagrams, data, images, or logos unless the user owns or is permitted to use them.
