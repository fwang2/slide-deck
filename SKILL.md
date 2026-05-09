---
name: feiyi-slide-deck
description: Generate styled PowerPoint decks from any input. Writes a self-contained python-pptx build script, runs it, validates the output, and delivers the .pptx.
---

# Slide Deck Generator

Generate professional PowerPoint presentations from conversational requests, markdown outlines, or raw documents. Produces a standalone `python-pptx` build script and the `.pptx` file.

## Style Config

Read the companion style config before generating any deck:

```
skills/feiyi-slide-deck/slide-deck-defaults.yaml
```

This YAML defines dimensions, fonts, font sizes, palette, card styling, speaker notes, and layout guardrails. All generated scripts must conform to these values unless the user provides overrides.

**Override mechanism:** The user can override any config value per-deck by:
- Passing a partial YAML file path
- Inline in their request (e.g., "use palette primary=00662C")
- Referencing a named preset file (e.g., `slide-deck-ornl.yaml` alongside the defaults)

Merge overrides onto defaults. Unspecified values keep their defaults.

## Workflow

Follow these steps in order:

### Step 1: Parse Input

Accept any input format:
- **Conversational:** "Build me a 10-slide deck about X covering Y and Z"
- **Markdown outline:** A structured list of slide titles and bullet points
- **Raw document:** Read the file, extract key content, identify themes and structure

If the user provides a `.pptx` template file, use it as the base presentation (inherit slide masters and layouts). Otherwise, build from blank layouts.

### Step 2: Propose Deck Structure

Present the user with a slide-by-slide outline before generating anything:

```
Proposed deck: "Title of Presentation" (N slides)

  1. [Title]    — Title of Presentation / Subtitle / Author
  2. [Content]  — Slide title — bullet summary of content
  3. [Content]  — Slide title — card grid: topic A, topic B, topic C
  ...
  N. [Closing]  — Key takeaways / contact info
```

Wait for the user to approve, rearrange, add, or remove slides. Do not proceed until approved.

### Step 3: Generate Build Script

Write a self-contained Python script that:
1. Defines style constants at the top (from config)
2. Defines helper functions inline (see Helper Functions section below)
3. Defines one function per slide (`slide_01_title()`, `slide_02_overview()`, etc.)
4. Calls all slide functions in order
5. Saves the `.pptx`
6. Runs OOXML validation and repair (see Validation section below)
7. Reports results

Save the script as `build_<deck-name>.py` in the working directory.

### Step 4: Run the Script

Execute the build script:

```bash
python3 build_<deck-name>.py
```

### Step 5: Report

Confirm to the user:
- Output file path
- Slide count
- Validation result (pass / warnings / errors)

## Slide Types

### Title Slide
- Full-width centered layout
- Title: heading font, 36-42pt, centered vertically in upper half
- Subtitle: body font, 18-20pt, below title
- Optional: date, author, affiliation in footer zone
- No cards, no slide chrome — clean and bold

### Content Slide
- **Top zone** (y=0.3 to y=1.3): slide title (28-32pt heading font, bold, heading color) + optional subtitle (16pt body font, muted color)
- **Content zone** (y=`layout.content_top` to y=`layout.content_bottom`): one of these arrangements:
  - **Bullets** — single or two-column, max `layout.max_bullets_per_slide` items, `font_sizes.body` pt minimum
  - **Cards** — 1 to `layout.max_cards_per_slide` rounded shadow cards with colored accent stripe/band, arranged as 1x1, 1x2, 1x3, 2x2, or 2x3 grid
  - **Code block** — dark background rectangle (`text.heading` color), code font, `font_sizes.code` pt minimum, with optional header label
  - **Mixed** — cards on one side, bullets or text on the other (60/40 or 50/50 split)
- **Footer zone** (below y=`layout.content_bottom`): optional thin rule + footer text (`font_sizes.footer` pt, muted color)

### Closing Slide
- Similar to title slide structure
- Supports a summary card or key takeaway band (accent card with `palette.primary`)
- Optional contact info / attribution in lower third

## Density & Balance Rules

Enforce these guardrails in every generated script:

1. **Never shrink fonts below minimums.** If content exceeds `max_bullets_per_slide` or `max_cards_per_slide` at minimum font sizes, split into multiple slides.
2. **Equal card sizing.** Cards in a grid must have equal width and consistent spacing (`layout.column_gap`).
3. **No overflow.** No text box or shape may extend below `layout.content_bottom`.
4. **Balanced layout.** Elements should be visually balanced — equal-width columns, consistent vertical spacing, content centered or evenly distributed rather than bunched to one side.

## Card Pattern

The card is the core visual element. Every generated script must implement cards exactly as described here.

### Anatomy

```
+-----+----------------------+
|     | LABEL TEXT           |   <- accent stripe (left) or band (top)
|  #  |                     |
|  #  | Body content here   |   <- white bg, rounded corners, drop shadow
|  #  | at 16pt minimum     |
|     |                     |
+-----+----------------------+
```

### Construction (python-pptx)

1. **Rounded rectangle** — `MSO_SHAPE.ROUNDED_RECTANGLE` with `adjustments[0] = cards.corner_radius`. Fill: `cards.background`. Line: none.
2. **Drop shadow** — Apply via OOXML on the shape's `spPr` element:
   ```python
   from pptx.oxml.ns import qn
   from lxml import etree

   def add_shadow(shape, offset_x_pt=1, offset_y_pt=2, blur_pt=4, opacity_pct=25):
       """Add a subtle drop shadow to a shape via OOXML."""
       spPr = shape._element.find(qn('a:spPr'))
       if spPr is None:
           spPr = shape._element.find('.//' + qn('p:spPr'))
       effectLst = spPr.find(qn('a:effectLst'))
       if effectLst is None:
           effectLst = etree.SubElement(spPr, qn('a:effectLst'))
       outerShdw = etree.SubElement(effectLst, qn('a:outerShdw'))
       outerShdw.set('blurRad', str(int(blur_pt * 12700)))
       outerShdw.set('dist', str(int(((offset_x_pt**2 + offset_y_pt**2)**0.5) * 12700)))
       outerShdw.set('dir', '2700000')  # ~SE direction
       outerShdw.set('algn', 'tl')
       outerShdw.set('rotWithShape', '0')
       srgbClr = etree.SubElement(outerShdw, qn('a:srgbClr'))
       srgbClr.set('val', '000000')
       alpha = etree.SubElement(srgbClr, qn('a:alpha'))
       alpha.set('val', str(opacity_pct * 1000))
   ```
3. **Accent stripe (left)** — `MSO_SHAPE.RECTANGLE` at the card's x, y, width = `cards.accent_width`, height = card height. Fill: accent color. Line: none.
4. **Accent band (top)** — Same but at card's x, y, width = card width, height = `cards.accent_width`. Fill: accent color. Line: none.
5. **Label text** — Textbox inside the card, bold, accent color, `font_sizes.label` pt, top of card interior.
6. **Body text** — Textbox below label, `text.body` color, `font_sizes.body` pt minimum.
7. **Padding** — Left: 0.15" (after stripe), top/bottom: 0.1", right: 0.1".

### Color Assignment

When multiple cards appear on one slide, assign accent colors by cycling through the palette in order:
`primary -> secondary -> accent1 -> accent2 -> accent3 -> primary -> ...`

Single-card slides use `palette.primary`.

### Code Blocks

Code blocks are a variant of cards:
- Dark background fill: `text.heading` color (default `1A1A2E`)
- No accent stripe
- Code font at `font_sizes.code` pt
- Light text: `text.on_accent` color
- Same rounded corners and shadow as regular cards

## Helper Functions

Every generated build script must define these helpers inline at the top, after the style constants. Adapt parameter defaults from the style config values.

### add_text

```python
def add_text(slide, x, y, w, h, text, *, size=16, bold=False, italic=False,
             color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=None):
    """Add a text box with word wrap and consistent margins."""
    if color is None:
        color = C_BODY
    if font is None:
        font = FONT_BODY
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb
```

### add_bullets

```python
def add_bullets(slide, x, y, w, h, items, *, size=16, color=None,
                bullet_color=None):
    """Add a bulleted list. Max items enforced by caller."""
    if color is None:
        color = C_BODY
    if bullet_color is None:
        bullet_color = PAL_PRIMARY
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(6)
        r1 = p.add_run()
        r1.text = "\u25b8  "
        r1.font.name = FONT_BODY
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = bullet_color
        r2 = p.add_run()
        r2.text = item
        r2.font.name = FONT_BODY
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
    return tb
```

### add_rounded_card

```python
def add_rounded_card(slide, x, y, w, h, label, body, accent_color=None,
                     accent_position="left"):
    """Add a rounded shadow card with accent stripe/band, label, and body text."""
    if accent_color is None:
        accent_color = PAL_PRIMARY

    # Card background
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.adjustments[0] = CARD_RADIUS
    shp.fill.solid()
    shp.fill.fore_color.rgb = C_CARD_BG
    shp.line.fill.background()
    add_shadow(shp)

    # Accent stripe or band
    if accent_position == "left":
        stripe = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, x, y, Inches(ACCENT_W), h)
        stripe.fill.solid()
        stripe.fill.fore_color.rgb = accent_color
        stripe.line.fill.background()
        text_x = x + Inches(ACCENT_W + 0.05)
        text_w = w - Inches(ACCENT_W + 0.15)
    else:  # top band
        band = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, x, y, w, Inches(ACCENT_W))
        band.fill.solid()
        band.fill.fore_color.rgb = accent_color
        band.line.fill.background()
        text_x = x + Inches(0.1)
        text_w = w - Inches(0.2)

    # Label
    label_y = y + Inches(0.1) if accent_position == "left" else y + Inches(ACCENT_W + 0.05)
    add_text(slide, text_x, label_y, text_w, Inches(0.4),
             label, size=FONT_SZ_LABEL, bold=True, color=accent_color)

    # Body
    body_y = label_y + Inches(0.35)
    body_h = h - (body_y - y) - Inches(0.1)
    add_text(slide, text_x, body_y, text_w, body_h,
             body, size=FONT_SZ_BODY, color=C_BODY)
```

### add_code_block

```python
def add_code_block(slide, x, y, w, h, code_text, header=None):
    """Add a dark-themed code block with optional header label."""
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.adjustments[0] = CARD_RADIUS
    shp.fill.solid()
    shp.fill.fore_color.rgb = C_HEADING
    shp.line.fill.background()
    add_shadow(shp)

    text_y = y + Inches(0.1)
    if header:
        add_text(slide, x + Inches(0.15), text_y, w - Inches(0.3), Inches(0.3),
                 header, size=FONT_SZ_LABEL, bold=True, color=C_ON_ACCENT,
                 font=FONT_CODE)
        text_y += Inches(0.35)

    code_h = h - (text_y - y) - Inches(0.1)
    add_text(slide, x + Inches(0.15), text_y, w - Inches(0.3), code_h,
             code_text, size=FONT_SZ_CODE, color=C_ON_ACCENT, font=FONT_CODE)
```

### slide_chrome

```python
def slide_chrome(slide, title, subtitle=None):
    """Add standard content-slide chrome: white background, title zone, footer rule."""
    # White background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG
    bg.line.fill.background()

    # Title
    add_text(slide, Inches(MARGIN), Inches(0.3),
             Inches(dimensions_w - 2 * MARGIN), Inches(0.8),
             title, size=FONT_SZ_TITLE, bold=True, color=C_HEADING,
             font=FONT_HEAD)

    # Subtitle
    if subtitle:
        add_text(slide, Inches(MARGIN), Inches(1.0),
                 Inches(dimensions_w - 2 * MARGIN), Inches(0.35),
                 subtitle, size=FONT_SZ_SUBTITLE, color=C_MUTED)

    # Footer rule
    rule = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(MARGIN), Inches(CONTENT_BOTTOM),
        Inches(dimensions_w - 2 * MARGIN), Emu(6000))
    rule.fill.solid()
    rule.fill.fore_color.rgb = C_MUTED
    rule.line.fill.background()
```

## Style Constants Preamble

Every generated build script must start with these imports and constants, derived from the style config:

```python
"""Build <deck-name>.pptx — generated by feiyi-slide-deck skill."""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
import zipfile
from io import BytesIO
from collections import Counter

# --- Style constants (from slide-deck-defaults.yaml) ---
# Dimensions
dimensions_w = 13.333
dimensions_h = 7.5

# Fonts
FONT_HEAD = "Aptos"
FONT_BODY = "Aptos"
FONT_CODE = "Fira Code"

# Font sizes (pt)
FONT_SZ_TITLE    = 28
FONT_SZ_SUBTITLE = 18
FONT_SZ_BODY     = 16
FONT_SZ_LABEL    = 12
FONT_SZ_CODE     = 12
FONT_SZ_FOOTER   = 10

# Palette
PAL_PRIMARY   = RGBColor(0x2B, 0x57, 0x9A)
PAL_SECONDARY = RGBColor(0x1A, 0x7F, 0x64)
PAL_ACCENT1   = RGBColor(0xC4, 0x65, 0x2A)
PAL_ACCENT2   = RGBColor(0x7B, 0x5E, 0xA7)
PAL_ACCENT3   = RGBColor(0xD4, 0xA8, 0x43)
PAL_NEUTRAL   = RGBColor(0x6B, 0x72, 0x80)
PALETTE_CYCLE = [PAL_PRIMARY, PAL_SECONDARY, PAL_ACCENT1, PAL_ACCENT2, PAL_ACCENT3]

# Text colors
C_HEADING  = RGBColor(0x1A, 0x1A, 0x2E)
C_BODY     = RGBColor(0x2D, 0x2D, 0x3A)
C_MUTED    = RGBColor(0x6B, 0x72, 0x80)
C_ON_ACCENT = RGBColor(0xFF, 0xFF, 0xFF)

# Card styling
C_CARD_BG   = RGBColor(0xFF, 0xFF, 0xFF)
CARD_RADIUS = 0.08
ACCENT_W    = 0.10  # inches

# Background
C_BG = RGBColor(0xFF, 0xFF, 0xFF)

# Layout
MARGIN         = 0.5
CONTENT_TOP    = 1.4
CONTENT_BOTTOM = 6.8
MAX_CARDS      = 6
MAX_BULLETS    = 7
COL_GAP        = 0.15

# Presentation setup
prs = Presentation()
prs.slide_width  = Inches(dimensions_w)
prs.slide_height = Inches(dimensions_h)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
```

When the user provides overrides, substitute the relevant constants. When a template `.pptx` is provided, replace `Presentation()` with `Presentation("path/to/template.pptx")`.

## OOXML Validation & Repair

Every generated build script must include this validation and repair code, called automatically after saving the `.pptx`:

```python
def fix_pptx(path):
    """Post-process the .pptx to fix known OOXML issues."""
    DML = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    with zipfile.ZipFile(path, 'r') as zin:
        buf = BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                    root = etree.fromstring(data)
                    # Fix DONUT shapes missing avLst
                    for pg in root.iter(f'{{{DML}}}prstGeom'):
                        if pg.get('prst') == 'donut':
                            avLst = pg.find(f'{{{DML}}}avLst')
                            if avLst is None:
                                avLst = etree.SubElement(pg, f'{{{DML}}}avLst')
                            if avLst.find(f'{{{DML}}}gd') is None:
                                gd = etree.SubElement(avLst, f'{{{DML}}}gd')
                                gd.set('name', 'adj')
                                gd.set('fmla', 'val 25000')
                    # Fix negative cx/cy
                    for xfrm in root.iter(f'{{{DML}}}xfrm'):
                        ext = xfrm.find(f'{{{DML}}}ext')
                        if ext is not None:
                            cx, cy = int(ext.get('cx', 0)), int(ext.get('cy', 0))
                            if cx < 0:
                                ext.set('cx', str(abs(cx)))
                                xfrm.set('flipH', '1')
                            if cy < 0:
                                ext.set('cy', str(abs(cy)))
                                xfrm.set('flipV', '1')
                    data = etree.tostring(root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
        buf.seek(0)
        with open(path, 'wb') as f:
            f.write(buf.read())


def validate_pptx(path):
    """Validate the .pptx for common OOXML issues. Returns list of error strings."""
    DML = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    PML = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    errors = []
    with zipfile.ZipFile(path) as z:
        if z.testzip():
            errors.append("Zip archive is corrupt")
            return errors
        for fname in z.namelist():
            if fname.startswith('ppt/slides/slide') and fname.endswith('.xml'):
                root = etree.fromstring(z.read(fname))
                for ext in root.iter(f'{{{DML}}}ext'):
                    if int(ext.get('cx', 0)) < 0 or int(ext.get('cy', 0)) < 0:
                        errors.append(f"{fname}: negative cx/cy")
                for pg in root.iter(f'{{{DML}}}prstGeom'):
                    if pg.get('prst') == 'donut':
                        avLst = pg.find(f'{{{DML}}}avLst')
                        if avLst is None or avLst.find(f'{{{DML}}}gd') is None:
                            errors.append(f"{fname}: DONUT missing avLst")
                ids = [el.get('id') for el in root.iter(f'{{{PML}}}cNvPr')]
                for sid, cnt in Counter(ids).items():
                    if cnt > 1:
                        errors.append(f"{fname}: duplicate cNvPr id={sid}")
    prs_check = Presentation(path)
    for i, slide in enumerate(prs_check.slides):
        for rId, rel in slide.part.rels.items():
            try:
                _ = rel.target_part
            except Exception:
                errors.append(f"Slide {i+1}: broken relationship {rId}")
    return errors
```

### Script Ending Pattern

Every generated build script must end with this pattern:

```python
# --- Build all slides ---
SLIDE_FUNCS = [
    slide_01_title,
    slide_02_xxx,
    # ... one entry per slide
    slide_NN_closing,
]

for fn in SLIDE_FUNCS:
    fn()

out = Path(__file__).parent / "<deck-name>.pptx"
prs.save(out)
fix_pptx(str(out))
errors = validate_pptx(str(out))
if errors:
    print(f"WARNING: {len(errors)} validation issue(s):")
    for e in errors:
        print(f"  - {e}")
else:
    print(f"Wrote {out}  ({len(SLIDE_FUNCS)} slides, validation PASS)")
```

## Speaker Notes

When `notes.enabled` is true in the config, inject speaker notes into every slide.

### Style: "bullets" (default)

Generate 3-5 key talking points as a bulleted list. Keep each point to one sentence.

### Style: "conversational"

Generate paragraph prose, ~140 wpm, calibrated to slide type:
- **Title slides:** 30-60 words (quick intro)
- **Content slides:** 150-280 words (lead with human problem before technical detail)
- **Closing slides:** 80-120 words (recap and call to action)

### Injection Method

Use the robust injection method that handles slides where `notes_text_frame` is null:

```python
NSMAP = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

def escape_xml(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def inject_notes(slide, text):
    """Inject speaker notes using robust XML method."""
    ns = slide.notes_slide
    if ns.notes_text_frame is not None:
        ns.notes_text_frame.text = text
        return
    sp_tree = ns._element.find('.//' + qn('p:spTree'))
    paras = "".join(
        f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{escape_xml(l)}</a:t></a:r></a:p>'
        if l.strip() else '<a:p/>'
        for l in text.split('\n')
    )
    sp_tree.append(etree.fromstring(
        f'<p:sp xmlns:p="{NSMAP["p"]}" xmlns:a="{NSMAP["a"]}">'
        f'<p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr><p:ph type="body" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr>'
        f'<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>'
    ))
```

Call `inject_notes(slide, notes_text)` at the end of each slide function.
