---
name: md-to-docx
description: Use when the user asks to convert a Chinese markdown document to .docx format, especially design docs, specs, or technical reports containing ASCII diagrams, tables, inline code, and bold formatting. Handles broken markdown fences and box-drawing characters common in Chinese-authored documents.
---

# Markdown to Word (Print-Ready)

Convert Chinese markdown documents to formatted .docx files suitable for printing and review. Handles the messy realities of Chinese-authored markdown: ASCII box-drawing diagrams outside code fences, broken ``` pairings, and inline formatting.

## When to Use

- User asks to "convert markdown to Word" or "output as docx" or "打印"
- Document contains Chinese text with mixed formatting
- Document has ASCII diagrams (┌┐└┘ characters) that render poorly in Word
- User wants a print-ready version without raw markup showing

## Core Script

Run `scripts/convert.py <input.md> [output.docx]`. The script is self-contained — copy it to any project and use it directly.

```
python scripts/convert.py design.md              # → design-print.docx
python scripts/convert.py design.md output.docx  # → output.docx
```

## What It Handles

| Markdown Source | Word Output |
|----------------|-------------|
| `# ## ### ####` headings | Heading 1/2/3 styles (SimHei/KaiTi) |
| `**bold text**` | Bold formatting |
| `` `inline code` `` | Consolas font, gray |
| Pipe tables `|...|` | Formatted tables (dark header, grid) |
| `- bullet` / `1. ordered` lists | Word List Bullet / List Number |
| `---` horizontal rules | Thin gray line |
| ``` fenced code blocks | Consolas monospace block |
| ASCII box diagrams (┌┐└┘) | Text description (auto-detected) |

## Three Critical Patterns

### 1. ASCII Diagram Detection

Chinese-authored markdown often has UTF-8 box-drawing characters **outside** ``` fences — they are raw text in the document:

```
┌──────────────────┐
│  Frontend (Vue3)  │
└────────┬─────────┘
```

**Detection:** Match characters in range `[─━│┃┌┐└┘├┤┬┴┼╭╮╰╯]`. Accumulate consecutive box-drawing lines into a buffer. Match against known diagram templates and replace with prose descriptions. Fallback: render as-is if no template matches.

### 2. Broken Code Fences

Many Chinese markdown documents have mismatched ``` pairs — a fence opens for an ASCII diagram but never closes. Everything after it gets swallowed as code.

**Fix:** Inside a code block, if a line matches `^#{1,4}\s` (heading pattern), auto-close the code block, flush collected content as a diagram (if it has box chars) or code, then process the heading normally.

```python
IS_HEADING = re.compile(r'^#{1,4}\s')
if in_code_block and IS_HEADING.match(line):
    flush_code_block_as_diagram_or_code()
    in_code_block = False
    # Fall through to process heading
```

### 3. Heading Priority

Heading detection must come **before** all other line processing. A line starting with `#` is always a heading — check this first, then check for tables, lists, boxes, etc.

```python
# Correct order:
if line.startswith('#'): process_heading(); continue
if is_table(line): process_table(); continue
if is_list(line): process_list(); continue
if is_box_drawing(line): accumulate_box(); continue
# ... normal text last
```

## Font Configuration

For Chinese documents, use actual Chinese font family names (not English aliases):

| Element | Font | Size |
|---------|------|------|
| Body | SimSun (宋体) | 12pt |
| Heading 1 | SimHei (黑体) | 22pt |
| Heading 2 | SimHei (黑体) | 16pt |
| Heading 3 | KaiTi (楷体) | 14pt |
| Inline code | Consolas | 10pt |
| Code block | Consolas | 9pt |

Set East-Asian font explicitly:
```python
run.font.name = 'SimSun'
run.element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
```

## Rich Text Parsing

Parse inline markdown by scanning for the next `**` or `` ` `` marker:

```python
def add_rich(paragraph, text):
    idx = 0
    while idx < len(text):
        bm = re.search(r'\*\*(.+?)\*\*', text[idx:])
        cm = re.search(r'`([^`]+?)`', text[idx:])
        # Find which marker comes first
        nxt = min(idx + bm.start() if bm else len(text),
                  idx + cm.start() if cm else len(text))
        if nxt > idx:
            paragraph.add_run(text[idx:nxt])
        # Apply formatting
        if bm and nxt == idx + bm.start():
            r = paragraph.add_run(bm.group(1)); r.bold = True
            idx = nxt + len(bm.group(0))
        elif cm and nxt == idx + cm.start():
            r = paragraph.add_run(cm.group(1))
            r.font.name = 'Consolas'; r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(80, 80, 80)
            idx = nxt + len(cm.group(0))
```

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Using English font names for Chinese | Use `SimSun`, `SimHei`, `KaiTi` and set `w:eastAsia` |
| Heading check after box detection | Move heading check BEFORE all other processing |
| rstrip() on heading line before checking | `rstrip()` only removes trailing whitespace — fine for `startswith('#')` |
| Forgetting to skip `|---|` table separator rows | Check with `re.match(r'^\|[\s\-:|]+\|$', line)` and skip |
| Not flushing box buffer before headings | Call `flush_box()` before checking for headings |
| Converting inside code block without heading guard | Add `IS_HEADING` check inside code block handler |

## Reference

- `scripts/convert.py` — the full reusable conversion script
