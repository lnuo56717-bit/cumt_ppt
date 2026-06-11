# Sample Workflow

This example assumes Codex has loaded the `cumt_ppt_mcp` MCP server.

## 1. Inspect a Deck

Call `inspect_ppt`:

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v1.pptx"
}
```

Then inspect a problem slide:

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v1.pptx",
  "slide_index": 14
}
```

## 2. Normalize CUMT Logo

Use slide 3 as the right-top logo reference and apply it from slide 4 onward:

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v1.pptx",
  "output_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v2_logo.pptx",
  "reference_slide_index": 3,
  "start_slide_index": 4,
  "skip_slide_indices": [1]
}
```

## 3. Generate a Three-Line Table

Use the experiment-setting slide:

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v2_logo.pptx",
  "output_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v3_table.pptx",
  "slide_index": 14,
  "table_data": [
    ["参数类别", "参数名称", "参数取值"],
    ["算法", "强化学习算法", "PPO"],
    ["工具", "训练工具", "Stable-Baselines3 / Gymnasium"],
    ["训练", "学习率", "3e-4"],
    ["训练", "折扣因子 γ", "0.99"]
  ],
  "left": 1.0,
  "top": 1.4,
  "width": 11.2,
  "height": 4.4
}
```

## 4. Apply Font Rules

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v3_table.pptx",
  "output_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v4_fonts.pptx",
  "title_font": "黑体",
  "body_font": "宋体",
  "latin_font": "Times New Roman"
}
```

## 5. Export Preview

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v4_fonts.pptx",
  "output_dir": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\preview_v4"
}
```

## 6. Check Quality

```json
{
  "pptx_path": "C:\\Users\\lenovo\\Documents\\pptskill\\out\\defense_v4_fonts.pptx",
  "expected_title": "滚筒智能生产线 AGV 自适应调度研究",
  "required_numbers": ["6", "953", "147", "31", "142", "22", "32428.0"]
}
```
