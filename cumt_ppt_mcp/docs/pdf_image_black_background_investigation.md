# PDF Image Black Background Investigation

## Problem

When figures are extracted from thesis PDFs, some images can appear with a black background even though they look white or transparent in the PDF reader. This affects PPT generation because extracted figures may look like dark blocks on a blue-white academic slide.

## Likely Causes

1. **Alpha channel compositing error**

   Some embedded PDF images contain transparency. If a tool extracts the raw image and saves it without compositing onto a white background, viewers or downstream converters may display transparent pixels as black.

2. **PDF page rendering with alpha enabled**

   Page renderers can produce transparent page backgrounds. If that rendered image is later opened by a tool that assumes black for transparent pixels, the slide image appears black.

3. **SVG/PNG conversion background fill**

   Vector figures or transparent PNGs can be rasterized with a default black background if no explicit background is supplied.

4. **Extraction method differences**

   - `PyMuPDF page.get_images()` extracts embedded image streams, but those streams may need alpha compositing.
   - `PyMuPDF page.get_pixmap(alpha=True)` can preserve transparent page background.
   - `PyMuPDF page.get_pixmap(alpha=False)` generally renders onto an opaque background.
   - `pdf2image` depends on Poppler availability and rendering flags.
   - PowerPoint/LibreOffice screenshots usually render what the viewer displays, but they are slower and less deterministic.

## Tested / Supported Strategies

### 1. PyMuPDF `get_images`

`extract_pdf_images_safe` extracts embedded images and converts them to RGBA. It then composites each image onto a caller-selected background, defaulting to white.

This addresses the most common black-background issue caused by alpha pixels.

### 2. PyMuPDF page render with white background

The tool also renders each page using `page.get_pixmap(alpha=False)` and saves the result after an explicit white-background pass. This gives a fallback visual reference when embedded figure extraction is incomplete or the PDF figure is vector-based.

### 3. pdf2image / Poppler

The tool reports whether Poppler appears available through `pdftoppm`. It does not require Poppler for the first implementation because adding a system dependency would make setup harder. A later version can add a `method="pdf2image"` option.

### 4. PowerPoint / LibreOffice screenshot method

This is useful for PPT preview export, but it is not the default for PDF figure extraction because it requires GUI or office software and is harder to run in a clean MCP environment.

## v0.3.0 Fix

The new MCP tool `extract_pdf_images_safe`:

- validates `.pdf` input paths;
- writes only to `output_dir`;
- extracts embedded images to `output_dir/images`;
- renders pages to `output_dir/page_renders`;
- composites transparent pixels onto `background_color`, default `white`;
- reports dark-pixel ratios for quick black-background screening;
- writes `image_inventory.json` and `extract_pdf_images_safe_report.md`.

## Recommendation

For thesis-PPT generation, use extracted images only when the figure is clear and visually appropriate. If an extracted embedded image looks wrong, use the rendered page crop or redraw the figure as editable PPT shapes/charts.
