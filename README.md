# feiyi-slide-deck

A Claude Code skill for generating professional PowerPoint presentations. Give it a topic, an outline, or a raw document — it writes a self-contained `python-pptx` build script, runs it, validates the OOXML output, and delivers the `.pptx`.

## Features

- **Any input format** — conversational request, markdown outline, or raw document
- **Self-contained scripts** — each generated build script is portable with no external dependencies beyond `python-pptx`
- **Default professional style** — clean white backgrounds, Aptos/Fira Code typography, rounded shadow cards with colored accent stripes
- **Style overrides** — swap palette, fonts, dimensions, or card styling per-deck via YAML or inline
- **Template support** — optionally start from an existing `.pptx` template
- **OOXML validation & repair** — automatically fixes known PowerPoint compatibility issues (DONUT avLst, negative dimensions, broken relationships)
- **Speaker notes** — bullet-point or conversational style, with robust XML injection

## Requirements

- [Claude Code](https://claude.ai/claude-code) (CLI, desktop app, or IDE extension)
- Python 3.8+ with `python-pptx` and `lxml`:
  ```bash
  pip install python-pptx lxml
  ```

## Installation

### Option 1: Clone into skills directory

```bash
cd ~/.claude/skills
git clone https://github.com/fwang2/slide-deck.git feiyi-slide-deck
```

### Option 2: Symlink from another location

```bash
git clone https://github.com/fwang2/slide-deck.git feiyi-slide-deck ~/projects/feiyi-slide-deck
ln -s ~/projects/feiyi-slide-deck ~/.claude/skills/feiyi-slide-deck
```

### Verify

Start a new Claude Code session. The skill should appear in your skill list. You can invoke it with:

```
/feiyi-slide-deck
```

Or just ask Claude to build you a presentation — it will detect and use the skill.

## Usage

### Conversational

> "Build me a 10-slide deck about distributed computing challenges in HPC"

### From an outline

> "Build a deck from this outline:
> 1. Title: Modern HPC Workflows
> 2. Three challenges (cards)
> 3. Current solutions (bullets)
> 4. Proposed architecture (cards)
> 5. Closing with key takeaways"

### From a document

> "Read raw-materials.docx and build a presentation from it"

### With overrides

> "Build the deck with palette primary=00662C, secondary=00B38F and accent position top"

### With a template

> "Use Facility-template/VISTA-talk.pptx as the template"

## Default Style

| Element | Value |
|---------|-------|
| Dimensions | 13.333" x 7.5" (16:9 widescreen) |
| Heading font | Aptos |
| Body font | Aptos |
| Code font | Fira Code |
| Body minimum | 16pt |
| Code minimum | 12pt |
| Card style | Rounded corners, drop shadow, colored left stripe |
| Background | White |

### Default Palette

| Name | Hex | Use |
|------|-----|-----|
| Primary | `#2B579A` | Titles, primary accent |
| Secondary | `#1A7F64` | Secondary cards |
| Accent 1 | `#C4652A` | Callouts, emphasis |
| Accent 2 | `#7B5EA7` | Variety |
| Accent 3 | `#D4A843` | Alerts, positive signals |
| Neutral | `#6B7280` | Borders, footers |

## Customization

Edit `slide-deck-defaults.yaml` to change the default style, or create additional preset files (e.g., `slide-deck-ornl.yaml`) alongside it.

## File Structure

```
feiyi-slide-deck/
  SKILL.md                    # Skill definition (workflow, patterns, code templates)
  slide-deck-defaults.yaml    # Default style configuration
  README.md                   # This file
```

## License

MIT
