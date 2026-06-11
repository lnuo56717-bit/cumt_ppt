# Sample Workflow

This example assumes Codex has loaded the `cumt_ppt_mcp` MCP server. Paths and
paper details below are fictional placeholders.

## 1. Inspect A Deck

Call `inspect_ppt`:

```json
{
  "pptx_path": "C:\\work\\demo_defense_v1.pptx"
}
```

Then inspect a problem slide:

```json
{
  "pptx_path": "C:\\work\\demo_defense_v1.pptx",
  "slide_index": 12
}
```

## 2. Inspect A Template

```json
{
  "template_pptx_path": "C:\\work\\reference_template.pptx"
}
```

Check `image_classifications` before applying the template. Large figures should
be classified as `content_image`, not `header_logo`.

## 3. Apply Template Style

```json
{
  "source_pptx_path": "C:\\work\\demo_defense_v1.pptx",
  "template_pptx_path": "C:\\work\\reference_template.pptx",
  "output_path": "C:\\work\\demo_defense_v2_template_style.pptx"
}
```

## 4. Normalize CUMT Logo

Use slide 3 as the right-top logo reference and apply it from slide 4 onward:

```json
{
  "pptx_path": "C:\\work\\demo_defense_v2_template_style.pptx",
  "output_path": "C:\\work\\demo_defense_v3_logo.pptx",
  "reference_slide_index": 3,
  "start_slide_index": 4,
  "skip_slide_indices": [1]
}
```

## 5. Generate A Three-Line Table

```json
{
  "pptx_path": "C:\\work\\demo_defense_v3_logo.pptx",
  "output_path": "C:\\work\\demo_defense_v4_table.pptx",
  "slide_index": 10,
  "table_data": [
    ["Category", "Parameter", "Value"],
    ["Algorithm", "Model", "ExampleNet"],
    ["Training", "Learning rate", "1e-3"],
    ["Evaluation", "Metric", "Accuracy"]
  ],
  "left": 1.0,
  "top": 1.4,
  "width": 11.2,
  "height": 4.4
}
```

## 6. Apply Font Rules

```json
{
  "pptx_path": "C:\\work\\demo_defense_v4_table.pptx",
  "output_path": "C:\\work\\demo_defense_v5_fonts.pptx",
  "title_font": "黑体",
  "body_font": "宋体",
  "latin_font": "Times New Roman"
}
```

## 7. Export Preview

```json
{
  "pptx_path": "C:\\work\\demo_defense_v5_fonts.pptx",
  "output_dir": "C:\\work\\preview_v5"
}
```

## 8. Check Quality

```json
{
  "pptx_path": "C:\\work\\demo_defense_v5_fonts.pptx",
  "expected_title": "Example Thesis Defense Title",
  "required_numbers": ["92.5", "0.01", "128"]
}
```
