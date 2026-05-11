---
name: ihg-brand
description: Apply the IHG Hotels & Resorts master brand to documents, decks, one pagers, briefings, and reports. Use this skill whenever the user says "make this on brand," "make it IHG branded," "use our brand," "use the master brand," "apply IHG styling," or asks for a Word doc, PowerPoint deck, one pager, briefing, report, memo, or charter where IHG branding should apply. Also use it whenever the deliverable is for an IHG audience and brand consistency matters, even if the user does not say the word "brand." Produces .docx and .pptx files with the IHG color palette, typography, footer, accent labels, and cover conventions while leaving full creative freedom for visual content (process flows, conceptual frameworks, custom diagrams, illustrative compositions).
---

# IHG Master Brand

Apply the IHG Hotels & Resorts master brand to PowerPoint decks and Word documents. The full color, typography, and layout specification lives in `references/brand_spec.md`. Read it once before producing anything for the first time in a conversation.

## Core philosophy

The brand is a **chrome and palette specification**, not a content template. The skill provides:

- Fixed colors, typography, and footer/header chrome that must stay consistent across every deliverable.
- Reusable cover, contents, and divider builders for routine page types.
- A small Python module (`scripts/ihg_brand.py`) exposing brand constants and chrome helpers.

Inside the chrome, **content design is open**. For visuals, diagrams, frameworks, process flows, quadrants, and any bespoke composition, exercise judgment. Compose shapes directly using `python-pptx`. Use the brand colors for fills and lines so the visual stays on brand, but do not try to force a custom diagram into one of the canned text layouts.

## Decide the format

- "Deck", "slides", "presentation", "PPT" -> PowerPoint (.pptx).
- "Doc", "Word", "report", "memo", "charter", "one pager", "briefing" -> Word (.docx).
- Ambiguous: ask, unless the existing content clearly fits one shape.

## Building a Word document

Read `references/brand_spec.md` once, then run `scripts/build_docx.py` with a JSON spec on stdin.

```bash
python3 scripts/build_docx.py /mnt/user-data/outputs/output.docx <<'JSON'
{
  "title": "AI Labs Charter",
  "subtitle": "Operating model and scope",
  "author": "Teresa Nguyen, AI Labs Lead",
  "doc_type": "report",
  "sections": [
    {"heading": "Purpose", "body": "AI Labs is..."},
    {"heading": "Scope", "body": "We operate as...",
     "subsections": [
       {"heading": "In scope", "bullets": ["Item one", "Item two"]},
       {"heading": "Out of scope", "bullets": ["Item three"]}
     ]}
  ]
}
JSON
```

`doc_type` options:
- `report` - full cover page + running header + footer with page numbers
- `memo` - compact header band with To/From/Date/Subject metadata block
- `briefing` - same as memo with audience field for distribution
- `one_pager` - tight margins, no cover, designed to fit on a single page

The script handles cover layout, headers, footers, heading hierarchy, and bullet styling automatically. Body content is yours to shape.

## Building a PowerPoint deck (the routine path)

For decks made mostly of text-driven slides (cover, contents, content lists, multi-column layouts, statements, dividers), use `scripts/build_pptx.py` with a JSON spec.

```bash
python3 scripts/build_pptx.py /mnt/user-data/outputs/output.pptx <<'JSON'
{
  "title": "Gemini Enterprise Rollout",
  "subtitle": "AI Labs production deployment update",
  "sections": [
    {"name": "Status", "slides": [
      {"layout": "title_only", "title": "Where we are"},
      {"layout": "content_text", "title": "Phase 1 milestones",
       "body": ["UAT complete", "Connectors live", "FinOps validated"]}
    ]}
  ]
}
JSON
```

Available layouts: `title_only`, `content_text`, `two_column`, `three_column`, `four_column`, `content_image`, `statement`, `section_divider`, `closing`, `three_wave`, `side_panel`, `logo_strip_cards`, `lifecycle_wheel`. See the docstring at the top of `build_pptx.py` for the full schema of each.

The pattern layouts handle the most common infographic shapes from the IHG visual vocabulary:
- `three_wave`: chevron ribbons over content cards. Use for evolution timelines, maturity stages, Excite to Activate to Expand to Transform style flows.
- `side_panel`: deep blue left rail with white text, content on the right. Use for chapter intros, pattern summaries, framed write-ups. Set `image_placeholder: true` to leave room for a diagram on the right.
- `logo_strip_cards`: row of up to five cards with colored header bands. Use for vendor comparisons, platform options, capability matrices.
- `lifecycle_wheel`: circular donut with 3 to 8 stages and an optional center label. Use for development lifecycles, governance loops, journey wheels.

The script auto-inserts a section divider before each section unless the first slide of the section is already a `section_divider` (so explicit dividers don't get duplicated). Set `"auto_section_dividers": false` at the spec root to suppress all auto dividers, or `"skip_divider": true` per section.

## Building a PowerPoint deck (the custom path)

When the deck includes process flows, conceptual frameworks, quadrants, custom diagrams, or any visual that doesn't fit the routine layouts, **do not stretch the JSON layouts**. Instead, write a custom Python script that imports `ihg_brand` and composes shapes directly. The brand chrome stays consistent; the content design is yours.

For common infographic patterns (circular lifecycles, swimlanes, stat callouts with arrows, buy/build option grids, vertical journey maps), read `references/visual_patterns.md` first. It has worked recipes that use the brand helpers correctly. The recipes are starting points, not templates: adjust geometry, color assignment, and label placement to fit the content.

### Pattern: hybrid build

When most of the deck is routine but a few slides need custom design, the cleanest pattern is:

1. Build the routine slides via the JSON spec, omitting the custom slides.
2. Open the resulting .pptx with `python-pptx`, insert the custom slides at the right positions, and save.

Alternatively, build the entire deck in a single custom Python script using `ihg_brand` helpers throughout. Use this when the deck is mostly bespoke.

### Custom slide template

```python
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from ihg_brand import (
    ORANGE, DEEP_BLUE, COOL_GRAY, INFO_COLORS, BLACK, WHITE,
    TEAL, LIGHT_TEAL, GOLD, SAND, MID_BLUE, PALE_BLUE,
    DEEP_BLUE_TINT, ORANGE_TINT, TEAL_TINT,
    SLIDE_W, SLIDE_H,
    new_slide, add_brand_chrome, add_slide_title, add_textbox,
    add_filled_rect, add_shape,
    # Pattern helpers (see references/visual_patterns.md for usage):
    add_card, add_chevron, add_numbered_circle, add_arrow_connector,
    add_donut_wedge, add_swimlane_header, add_section_badge, add_star_marker,
)

prs = Presentation("path/to/deck-from-routine-build.pptx")

slide = new_slide(prs)
add_brand_chrome(slide, page_num=5, section_name="STRATEGY")
add_slide_title(slide, "Build vs buy framework")

# ... your custom shapes here ...

prs.save("path/to/final-deck.pptx")
```

### Brand colors for custom visuals

- **Primary accents**: `ORANGE` for highlights, calls to action, the "current state" or "selected" element. Use sparingly so it stays an accent.
- **Secondary structure**: `DEEP_BLUE` for headers, important labels, the second emphasis tier.
- **Quiet structure**: `COOL_GRAY` for grid lines, dividers, deemphasized elements, and shape outlines when filled shapes would compete with content.
- **Charts and series**: `INFO_COLORS` (six colors: teal, light teal, gold, sand, mid blue, pale blue). Use in order. Reserve `ORANGE` for the highlighted series only; do not include it in a normal chart palette.
- **Backgrounds**: `WHITE` for page; `WARM_WHITE` only for the footer band (already handled by `add_footer`).

### Worked example: process flow

```python
slide = new_slide(prs)
add_brand_chrome(slide, page_num=4, section_name="PROCESS")
add_slide_title(slide, "Engagement model")

stages = ["Intake", "Scope", "Build", "Evaluate", "Handoff"]
n = len(stages)
total_w = 11.5
gap = 0.2
box_w = (total_w - gap * (n - 1)) / n
box_h = 1.2
top = 3.2
left_start = 0.6

for i, stage in enumerate(stages):
    x = Inches(left_start + i * (box_w + gap))
    fill = ORANGE if i == 0 else DEEP_BLUE
    box = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                    x, Inches(top), Inches(box_w), Inches(box_h),
                    fill=fill)
    add_textbox(slide, x, Inches(top), Inches(box_w), Inches(box_h),
                stage, size=16, bold=True, color=WHITE,
                align=1, anchor=3)  # PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE

    # Arrow between boxes
    if i < n - 1:
        ax = Inches(left_start + (i + 1) * box_w + i * gap)
        add_shape(slide, MSO_SHAPE.RIGHT_ARROW,
                  ax, Inches(top + 0.5),
                  Inches(gap), Inches(0.2),
                  fill=COOL_GRAY)
```

### Worked example: 2x2 quadrant

```python
slide = new_slide(prs)
add_brand_chrome(slide, page_num=6, section_name="STRATEGY")
add_slide_title(slide, "Capability vs effort")

# Frame
left, top = Inches(2.5), Inches(2.4)
size = Inches(4.5)
add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, size, size,
          fill=None, line_color=COOL_GRAY)
# Cross lines
mid_x = left + size / 2
mid_y = top + size / 2
add_shape(slide, MSO_SHAPE.RECTANGLE, mid_x, top, Emu(0), size,
          fill=None, line_color=COOL_GRAY)  # vertical
# (use add_shape with MSO_SHAPE.LINE for proper lines, omitted for brevity)

# Quadrant labels
quadrants = [
    ("High value, low effort", "QUICK WINS", ORANGE),
    ("High value, high effort", "BIG BETS", DEEP_BLUE),
    ("Low value, low effort", "FILLERS", COOL_GRAY),
    ("Low value, high effort", "AVOID", BLACK),
]
# ... position each in its quadrant ...
```

The point of these examples is the pattern, not the exact code: import the brand module, call `new_slide` + `add_brand_chrome` so the chrome is right, then use `add_shape`, `add_filled_rect`, and `add_textbox` to compose freely.

## Brand colors and fonts (quick reference)

For full spec see `references/brand_spec.md`. Custom-visual usage hierarchy is in `references/visual_patterns.md`.

Primary palette:

| Color            | Hex      | Use                                             |
| ---------------- | -------- | ----------------------------------------------- |
| Signature Orange | #E8542C  | Subtitles, accent labels, highlight one element |
| Deep Blue        | #1F4456  | Secondary headers, structured content           |
| Warm White       | #F0EEED  | Footer band background                          |
| Cool Gray        | #C2C7CA  | Dividers, deemphasized lines                    |
| Black            | #262626  | Body text, titles (note: not pure black)        |
| Rewards Gray     | #6F8996  | ONLY when content features IHG Rewards          |

Infographics palette (named constants in `ihg_brand`):

| Constant   | Hex      | Use                                                       |
| ---------- | -------- | --------------------------------------------------------- |
| TEAL       | #448790  | Primary chart series, secondary structural fill           |
| LIGHT_TEAL | #A1C6C8  | Light variant: card headers needing softer weight         |
| GOLD       | #F4C064  | Accent fills (renders best with black text)               |
| SAND       | #EDD8BC  | Quaternary series, tonal card fills                       |
| MID_BLUE   | #426399  | Mid-tier headers, card variation                          |
| PALE_BLUE  | #D9E5EC  | Background fills, sixth chart series                      |

Tonal tints (for card body fills, alternating rows):

| Constant         | Hex      | Use                                                |
| ---------------- | -------- | -------------------------------------------------- |
| DEEP_BLUE_TINT   | #E6ECEF  | Card body under deep blue header, soft row band    |
| ORANGE_TINT      | #FCE8E1  | Highlight band for "current" or "recommended"      |
| TEAL_TINT        | #E3EEEE  | Soft alternating row, card body                    |

Typography is Calibri throughout. Titles use Light weight, body uses Regular.

## Things to never do

- Never substitute the colors with "close enough" values. Use the constants in `ihg_brand`.
- Never use `REWARDS_GRAY` unless the deliverable explicitly features IHG Rewards.
- Never omit the footer on internal materials.
- Never claim the document is "approved by Brand."

## After generating

Always call `present_files` with the output path so the user can download it. Briefly say what was produced. Do not narrate the contents back at length.