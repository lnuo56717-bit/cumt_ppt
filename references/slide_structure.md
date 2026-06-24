# Recommended Slide Structure

## Journal / Conference Paper Presentation (12–16 slides)

Use for APL, Nature, IEEE, or similar peer-reviewed papers. Adapt slide count to paper complexity.

1. **Cover** — title, authors + affiliations, journal name + DOI, presentation date
2. **Outline** — six-section structure; doubles as breadcrumb reference
3. **Research Background & Motivation** — problem statement, literature gap, significance
4. **Device / Method Design** — structure diagram, design parameters, materials
5. **Working Principle** — physics mechanism, resonance condition, key formula
6. **Simulation Analysis** — FDTD / FEM results, absorption spectra, field distributions
7. **Fabrication & Characterization** — process flow, microscopy images, R-T / I-V curves
8. **Experimental Setup** — measurement system diagram, light source, detector configuration
9. **Key Results** — absorption or response spectra, peak values, wavelength selectivity
10. **Performance Metrics** — NEP, D*, responsivity table; comparison with prior art
11. **Conclusions & Outlook** — summary of contributions, limitations, future work
12. **Acknowledgements / Q&A** — funding sources, collaborators, thank-you

### Journal Breadcrumb Sections (example for photonics detector paper)

```python
SECTIONS = [
    ("01", "研究背景"),
    ("02", "器件设计"),
    ("03", "仿真分析"),
    ("04", "制备实验"),
    ("05", "结果性能"),
    ("06", "结论展望"),
]
# Slide → 0-based section index; None for cover/outline/thank-you
SLIDE_SECTION = {
    1: None,   # Cover
    2: None,   # Outline
    3: 0,      # Background
    4: 1,      # Design
    5: 1,      # Working Principle
    6: 2,      # Simulation
    7: 3,      # Fabrication
    8: 3,      # Exp Setup
    9: 4,      # Results
    10: 4,     # Performance
    11: 4,     # Metrics Table
    12: 5,     # Conclusions
    13: None,  # Thank you
}
```

---

## Thesis Defense (14–20 slides)

Use 14-20 slides, adjusting only when the thesis scope or user request requires it.

1. 封面
2. 汇报提纲
3. 研究背景
4. 研究意义与核心问题
5. 企业或场景概况
6. 车间布局与物流路径
7. 路径冲突场景
8. 死锁或问题机理
9. 研究内容与技术路线
10. 调度问题建模
11. 算法策略设计
12. 安全约束或改进机制
13. 混合调度框架
14. 仿真环境与实验设置
15. 算法对比结果
16. 消融实验分析
17. 多尺度或扩展实验分析
18. 研究结论
19. 不足与展望
20. 致谢

## Notes

- Combine adjacent pages only if the deck is too long and the combined slide remains readable.
- Do not force pages with weak content.
- Keep required pages when relevant: outline, method/framework, experiment settings, experiment results, conclusion/future work, thanks.
