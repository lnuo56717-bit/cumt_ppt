# Prompt: Generate PPT from Journal Paper PDF

Use this prompt when the source is a journal or conference paper (APL, Nature, IEEE, etc.), not a CUMT thesis.

---

## Inputs required from the user

- **Paper PDF path**: the journal paper to present.
- **Output PPT path**: where to save the generated deck.
- **Figures folder** (optional): a folder of extracted/curated figures. If absent, extract from PDF automatically.
- **Reference template PPT** (optional): for layout and color style only; do NOT copy its content.
- **Slide count target** (optional, default: 12–13 slides).

---

## Workflow

### Step 1 — Extract content from PDF

Read the full PDF text. Extract:
- Title, authors, affiliations, journal name, volume, DOI, submission/acceptance dates.
- Abstract: background, gap, contribution, result summary.
- Introduction: motivation, prior art references, key claims.
- Methods/Device section: structure, materials, parameters, design equations.
- Simulation section: software (FDTD/FEM), key results, figures.
- Experiment section: fabrication process, characterization setup, measurement equipment.
- Results section: spectra, response curves, tables, peak values, error bars.
- Performance comparison table (if present).
- Conclusions and future work.

### Step 2 — Extract figures from PDF

Run `extract_pdf_images_safe` MCP tool on the PDF, or use PyMuPDF:
```python
import fitz
doc = fitz.open(pdf_path)
for page_num, page in enumerate(doc):
    for img_info in page.get_images(full=True):
        xref = img_info[0]
        base_img = doc.extract_image(xref)
        # save base_img["image"] with extension base_img["ext"]
```

Create a figure inventory: page, index, pixel size, likely caption label, recommended slide.
Skip images smaller than 200 px in either dimension (icons, logos, etc.).

### Step 3 — Render formulas

For every key equation in the paper, call `render_formula_png` MCP tool, or run:
```bash
python scripts/render_formulas.py --out <formulas_dir>
```

Rules:
- One `$...$` block per render call; merge adjacent math blocks.
- Do NOT use `\&`; use `\mathrm{and}` instead.
- Typical formulas to render: resonance condition, efficiency metric, NEP, D*, responsivity.

### Step 4 — Plan slide structure

Map paper sections to slides using the journal spine in `references/slide_structure.md`.

Define SECTIONS and SLIDE_SECTION constants:
```python
SECTIONS = [("01","研究背景"),("02","器件设计"),("03","仿真分析"),
            ("04","制备实验"),("05","结果性能"),("06","结论展望")]

SLIDE_SECTION = {
    1: None,   # Cover
    2: None,   # Outline
    3: 0,      # Background
    # ... map remaining slides to 0-based section indices
    -1: None,  # Thank you / Q&A
}
```

### Step 5 — Build the deck with python-pptx

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
TITLE_H = Inches(1.0)
CRUMB_H = Inches(0.32)
CONTENT_T = TITLE_H + CRUMB_H   # 1.32" from top

DEEP_BLUE = RGBColor(0x1F, 0x49, 0x7D)
GOLD      = RGBColor(0xFF, 0xC0, 0x00)
```

For every slide:
1. Call `title_bar(slide, title, slide_num)` — blue bar at top with number badge.
2. Call `add_breadcrumb_strip` MCP tool or draw rectangle tabs manually.
3. Add content below `CONTENT_T`.

### Step 6 — Font rules (critical)

```python
from lxml import etree
from pptx.oxml.ns import qn

def set_font(run, name_cn=None, name_en="Times New Roman",
             size=None, bold=False, color=None, italic=False):
    rf = run.font
    if name_en:
        rf.name = name_en         # sets Latin font
    if name_cn:
        rPr = run._r.get_or_add_rPr()
        for tag, face in [('a:latin', name_en or '+mn-lt'),
                          ('a:ea', name_cn), ('a:cs', name_cn)]:
            el = rPr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rPr, qn(tag))
            el.set('typeface', face)
    if size: rf.size = Pt(size)
    rf.bold = bold; rf.italic = italic
    if color: rf.color.rgb = color
```

DO NOT do `rf.name = name_cn` then `rf.name = name_en` — the second assignment overwrites the first and the Chinese font is lost. Always set East Asian font via XML `a:ea`/`a:cs`.

### Step 7 — Insert figures and formulas

```python
from pptx.util import Inches
from PIL import Image

def insert_img(slide, path, l, t, w):
    im = Image.open(path)
    h = int(w * im.height / im.width)
    return slide.shapes.add_picture(path, l, t, w, h)
```

Place figures on the slide relevant to their paper section. Do not place a simulation figure on an experiment slide.

### Step 8 — Common Chinese string literal pitfall

Do NOT put curly/smart quotes inside Python double-quoted strings:
```python
# BAD  — "（" triggers SyntaxError in some editors
text = "光谱范围为1025 cm⁻¹（"long wave IR）"

# GOOD — use ASCII quotes or square brackets
text = "光谱范围为1025 cm^-1 [long wave IR]"
```

### Step 9 — Validate

- Run `check_ppt_quality` MCP tool.
- Check that every content slide has a breadcrumb strip with the correct active section.
- Check that all inserted images are readable and on-topic.
- Check that formula images are not blurry (dpi >= 160 recommended).

---

## Output

- Primary: `<output_path>.pptx`
- Optional formula images: `<formulas_dir>/*.png`
- Optional figure extracts: `<figures_dir>/*.jpeg` or `*.png`
- Validation report: printed to console or saved alongside the PPTX.
