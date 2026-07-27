#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convert Chinese markdown to print-ready .docx.

Usage:
    python convert.py input.md                  # → input-print.docx
    python convert.py input.md output.docx      # → output.docx

Handles:
    - Headings (# ## ### ####)
    - Bold (**text**) and inline code (`code`)
    - Pipe tables with header formatting
    - Bullet and numbered lists
    - Fenced code blocks (```)
    - ASCII box-drawing diagrams → text descriptions (auto-detected)
    - Broken ``` fences (auto-close on heading detection)
"""
import re, sys, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Configuration ──────────────────────────────────────────────
FONT_BODY   = 'SimSun'
FONT_HEI    = 'SimHei'
FONT_KAI    = 'KaiTi'
FONT_CODE   = 'Consolas'
SIZE_BODY   = Pt(12)
SIZE_H1     = Pt(22)
SIZE_H2     = Pt(16)
SIZE_H3     = Pt(14)
SIZE_CODE   = Pt(10)
SIZE_CBLOCK = Pt(9)
COLOR_CODE  = RGBColor(80, 80, 80)
COLOR_CBLK  = RGBColor(51, 51, 51)
COLOR_DESC  = RGBColor(40, 40, 40)
COLOR_HR    = RGBColor(180, 180, 180)

BOX_RE = re.compile(r'[─━│┃┌┐└┘├┤┬┴┼╭╮╰╯▁▔]')
HEADING_RE = re.compile(r'^#{1,4}\s')

# ── Rich text parser ───────────────────────────────────────────
def add_rich(paragraph, text):
    """Parse **bold** and `code` and add as Word runs."""
    idx = 0
    while idx < len(text):
        bm = re.search(r'\*\*(.+?)\*\*', text[idx:])
        cm = re.search(r'`([^`]+?)`', text[idx:])
        nxt = len(text)
        if bm and idx + bm.start() < nxt: nxt = idx + bm.start()
        if cm and idx + cm.start() < nxt: nxt = idx + cm.start()
        if nxt > idx: paragraph.add_run(text[idx:nxt])
        if nxt >= len(text): break
        if bm and idx + bm.start() == nxt:
            r = paragraph.add_run(bm.group(1)); r.bold = True
            idx = nxt + len(bm.group(0))
        elif cm and idx + cm.start() == nxt:
            r = paragraph.add_run(cm.group(1))
            r.font.name = FONT_CODE; r.font.size = SIZE_CODE
            r.font.color.rgb = COLOR_CODE
            idx = nxt + len(cm.group(0))
        else:
            paragraph.add_run(text[idx]); idx += 1

# ── Document builders ──────────────────────────────────────────
def build_doc():
    doc = Document()
    ns = doc.styles['Normal']
    ns.font.name = FONT_BODY; ns.font.size = SIZE_BODY
    ns.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_BODY)
    for i, (fn, fs) in enumerate([(FONT_HEI, SIZE_H1), (FONT_HEI, SIZE_H2), (FONT_KAI, SIZE_H3)]):
        h = doc.styles['Heading %d' % (i+1)]; f = h.font
        f.name = fn; f.size = fs
        h.element.rPr.rFonts.set(qn('w:eastAsia'), fn)
    return doc

def add_p(doc, text):
    p = doc.add_paragraph(); add_rich(p, text)

def add_li(doc, text, num=False):
    p = doc.add_paragraph(style='List Number' if num else 'List Bullet')
    p.clear(); add_rich(p, text)

def add_table(doc, rows):
    nc = max(len(r) for r in rows)
    t = doc.add_table(rows=len(rows), cols=nc, style='Table Grid')
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, rd in enumerate(rows):
        for j, ct in enumerate(rd):
            if j < nc:
                c = t.cell(i, j); c.text = ''
                pc = c.paragraphs[0]; add_rich(pc, ct)
                if i == 0:
                    for r in pc.runs: r.bold = True; r.font.size = Pt(10)
                    sh = OxmlElement('w:shd')
                    sh.set(qn('w:fill'), '1F4E79'); sh.set(qn('w:val'), 'clear')
                    c._tc.get_or_add_tcPr().append(sh)
                    for r in pc.runs: r.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    for r in pc.runs: r.font.size = Pt(10)

def add_code(doc, lines):
    for ln in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.left_indent = Cm(1)
        r = p.add_run(ln)
        r.font.name = FONT_CODE; r.font.size = SIZE_CBLOCK
        r.font.color.rgb = COLOR_CBLK

def add_desc(doc, text):
    for ln in text.strip().split('\n'):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5); add_rich(p, ln)
        for r in p.runs:
            r.font.size = Pt(10.5)
            if not r.bold and r.font.name != FONT_CODE:
                r.font.color.rgb = COLOR_DESC

def add_hr(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run('─' * 50); r.font.color.rgb = COLOR_HR

# ── Diagram → text descriptions ────────────────────────────────
# Extend this dict with your own diagram templates.
# Key: a substring unique to the diagram.  Value: prose description.
DIAGRAM_TEMPLATES = {
    'REST API': (
        '[系统部署架构]\n四层: 前端(Vue3+ElementPlus) -> REST API -> 后端(FastAPI, 基础模块复用+业务模块新建) -> 基础设施(RAGFlow/MinIO/Redis/PostgreSQL+pgvector/Ollama/BGE等)。'
    ),
    'Schema: hr': (
        '[数据架构: 同库不同Schema]\nhr Schema(政工写): personnel/department/position/leave_log/transfer_log。\npolice_admin Schema(内勤写): tasks/journals/schedules/sops/skills/knowledge_bases等。内勤对hr仅SELECT。\npublic Schema: 系统表。'
    ),
    '全局共享库': (
        '[知识库体系]\n全局共享库(局办维护, 所有单位只读) + 各单位私有库(自动归属, 可选择性共享给指定单位或全局, 可撤销)。'
    ),
}

def match_diagram(raw):
    """Return prose description if the diagram matches a known template."""
    for key, desc in DIAGRAM_TEMPLATES.items():
        if key in raw:
            return desc
    return None

# ── Main parser ────────────────────────────────────────────────
def convert(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    doc = build_doc()
    i = 0; in_cb = False; cb = []; tb = []; box_buf = []

    def flush_box():
        nonlocal box_buf
        if box_buf:
            raw = '\n'.join(box_buf)
            desc = match_diagram(raw)
            if desc: add_desc(doc, desc)
            box_buf = []

    while i < len(lines):
        line = lines[i].rstrip()

        # Skip YAML frontmatter
        if line == '---' and i == 0:
            i += 1
            while i < len(lines) and lines[i].strip() != '---': i += 1
            i += 1; continue

        # Code fence
        if line.startswith('```'):
            if in_cb:
                raw = '\n'.join(cb)
                desc = match_diagram(raw)
                if desc: add_desc(doc, desc)
                else: add_code(doc, cb)
                cb = []; in_cb = False
            else:
                flush_box(); in_cb = True
            i += 1; continue

        # Inside code block — auto-close if heading detected (broken fence)
        if in_cb:
            if HEADING_RE.match(line):
                raw = '\n'.join(cb)
                desc = match_diagram(raw)
                if desc: add_desc(doc, desc)
                elif cb: add_code(doc, cb)
                cb = []; in_cb = False
                # Fall through to process this heading
            else:
                cb.append(line); i += 1; continue

        # ASCII box-drawing (standalone, no code fence)
        if BOX_RE.search(line) and not line.startswith('|'):
            flush_box(); box_buf.append(line); i += 1; continue
        if box_buf and (not line.strip() or re.match(r'^[\s│┃├┤┬┴┼╭╮╰╯]+', line)):
            box_buf.append(line); i += 1; continue
        if box_buf and BOX_RE.search(line):
            box_buf.append(line); i += 1; continue
        flush_box()

        # Headings (highest priority)
        if line.startswith('#### '): doc.add_heading(line[5:], level=3); i += 1; continue
        if line.startswith('### '):  doc.add_heading(line[4:], level=3); i += 1; continue
        if line.startswith('## '):   doc.add_heading(line[3:], level=2); i += 1; continue
        if line.startswith('# '):    doc.add_heading(line[2:], level=1); i += 1; continue

        # Table
        if line.startswith('|') and line.endswith('|'):
            tb.append(line); i += 1
            while i < len(lines) and lines[i].strip().startswith('|') and lines[i].strip().endswith('|'):
                tb.append(lines[i].rstrip()); i += 1
            rows = []
            for tl in tb:
                tl = tl.strip()
                if re.match(r'^\|[\s\-:|]+\|$', tl): continue
                rows.append([c.strip() for c in tl.split('|')[1:-1]])
            if rows: add_table(doc, rows)
            tb = []; continue

        # Horizontal rule
        if line.strip() == '---':
            add_hr(doc); i += 1; continue

        # Lists
        if re.match(r'^[\s]*\- ', line):
            add_li(doc, re.sub(r'^[\s]*\- ', '', line)); i += 1; continue
        if re.match(r'^\d+\.\s', line):
            add_li(doc, re.sub(r'^\d+\.\s', '', line), True); i += 1; continue

        # Skip orphan box-drawing lines
        if re.match(r'^[\s]*[┌┐└┘├┤│─┬┴┼╭╮╰╯▁▔]', line):
            i += 1; continue

        # Normal paragraph
        if line.strip():
            add_p(doc, line)

        i += 1

    # Final cleanup
    flush_box()
    if in_cb and cb:
        raw = '\n'.join(cb)
        desc = match_diagram(raw)
        if desc: add_desc(doc, desc)
        else: add_code(doc, cb)

    for sec in doc.sections:
        sec.top_margin = Cm(2.54); sec.bottom_margin = Cm(2.54)
        sec.left_margin = Cm(3.18); sec.right_margin = Cm(3.18)

    doc.save(output_path)
    return output_path

# ── CLI ────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python convert.py input.md [output.docx]')
        sys.exit(1)
    inp = sys.argv[1]
    if len(sys.argv) >= 3:
        out = sys.argv[2]
    else:
        base = os.path.splitext(inp)[0]
        out = base + '-print.docx'
    result = convert(inp, out)
    print('Done: ' + result)
