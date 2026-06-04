# Example Input Contract

This is a fictional example. Do not store real thesis PDFs, PPTs, templates, or private figures in the skill.

```yaml
task: generate_or_polish_cumt_defense_ppt
mode: generate_from_pdf
thesis_pdf: C:\Users\student\Documents\fictional_thesis.pdf
figure_folder: C:\Users\student\Documents\fictional_figures
existing_ppt: null
reference_template: C:\Users\student\Documents\optional_style_template.pptx
output_ppt: C:\Users\student\Documents\outputs\fictional_defense_v1.pptx
preview_folder: C:\Users\student\Documents\outputs\preview
title_exact: 基于智能优化的矿山物流调度方法研究
student_name: 张三
advisor: 李四 副教授
college: 矿业工程学院
major: 工业工程
date: 2026年6月
experiment_data:
  baseline:
    completed_tasks: 10
    conflict_events: 120
    total_score: -3500.0
  proposed:
    completed_tasks: 88
    conflict_events: 5
    total_score: 18600.0
constraints:
  - do_not_change_title
  - do_not_invent_data
  - do_not_copy_template_content
  - match_images_to_slide_topics
```
