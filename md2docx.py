#!/usr/bin/env python3
"""
MD to DOCX Converter by Jair Lima
Converts Markdown files to Word DOCX with proper formatting.

Usage:
  md2docx                    # Convert all .md in current folder
  md2docx arquivo.md         # Convert specific file
  md2docx --force arquivo.md # Overwrite even if DOCX already exists
"""

import sys
import os
import re
import argparse
from pathlib import Path

# Fix Windows terminal encoding
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import mistune
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT


# ─────────────────────────────────────────────────────────────────────────────
# Document styles setup
# ─────────────────────────────────────────────────────────────────────────────

def setup_styles(doc: Document):
    """Configure all document styles for proper DOCX output."""

    # ── Normal (base) ────────────────────────────────────────────────────────
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    normal.paragraph_format.line_spacing = 1.15

    # ── Headings ─────────────────────────────────────────────────────────────
    heading_config = [
        ("Heading 1", 22, True, RGBColor(0x1F, 0x39, 0x64), Pt(12), Pt(6)),
        ("Heading 2", 16, True, RGBColor(0x2E, 0x74, 0xB5), Pt(10), Pt(4)),
        ("Heading 3", 13, True, RGBColor(0x1F, 0x39, 0x64), Pt(8),  Pt(4)),
        ("Heading 4", 12, True, RGBColor(0x2E, 0x74, 0xB5), Pt(6),  Pt(2)),
        ("Heading 5", 11, True, RGBColor(0x40, 0x40, 0x40), Pt(4),  Pt(2)),
        ("Heading 6", 11, False, RGBColor(0x70, 0x70, 0x70), Pt(4), Pt(2)),
    ]
    for name, size, bold, color, space_before, space_after in heading_config:
        style = doc.styles[name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.color.rgb = color
        style.paragraph_format.space_before = space_before
        style.paragraph_format.space_after = space_after
        style.paragraph_format.keep_with_next = True

    # ── Code (inline/block character style) ──────────────────────────────────
    if "Code Char" not in [s.name for s in doc.styles]:
        code_char = doc.styles.add_style("Code Char", 2)  # 2 = character style
    else:
        code_char = doc.styles["Code Char"]
    code_char.font.name = "Courier New"
    code_char.font.size = Pt(10)

    # ── Code Block (paragraph) ────────────────────────────────────────────────
    if "Code Block" not in [s.name for s in doc.styles]:
        code_block = doc.styles.add_style("Code Block", 1)
        code_block.base_style = doc.styles["Normal"]
    else:
        code_block = doc.styles["Code Block"]
    code_block.font.name = "Courier New"
    code_block.font.size = Pt(9.5)
    code_block.paragraph_format.space_before = Pt(6)
    code_block.paragraph_format.space_after = Pt(6)
    code_block.paragraph_format.left_indent = Cm(0.5)
    code_block.paragraph_format.right_indent = Cm(0.5)

    # ── Block Quote ───────────────────────────────────────────────────────────
    if "Block Quote" not in [s.name for s in doc.styles]:
        bq = doc.styles.add_style("Block Quote", 1)
        bq.base_style = doc.styles["Normal"]
    else:
        bq = doc.styles["Block Quote"]
    bq.font.name = "Calibri"
    bq.font.size = Pt(11)
    bq.font.italic = True
    bq.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    bq.paragraph_format.left_indent = Cm(1.0)
    bq.paragraph_format.right_indent = Cm(1.0)
    bq.paragraph_format.space_before = Pt(4)
    bq.paragraph_format.space_after = Pt(4)


def set_paragraph_shading(para, fill_hex: str):
    """Add background shading to a paragraph."""
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    pPr.append(shd)


def set_left_border(para, color_hex: str = "2E74B5", size: int = 24):
    """Add a left border to a paragraph (blockquote style)."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), color_hex)
    pBdr.append(left)
    pPr.append(pBdr)


def add_hyperlink(para, text: str, url: str):
    """Add a hyperlink run to a paragraph."""
    part = para.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    r.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    hyperlink.append(r)
    para._p.append(hyperlink)
    return hyperlink


def add_horizontal_rule(doc: Document):
    """Add a horizontal rule paragraph."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "AAAAAA")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return para


def set_table_style(table):
    """Apply clean table borders and shading."""
    tbl = table._tbl
    tblPr = tbl.tblPr

    # Table borders
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "AAAAAA")
        tblBorders.append(border)
    tblPr.append(tblBorders)


def shade_table_header(row):
    """Apply header shading to a table row."""
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "DEEAF1")
        tcPr.append(shd)
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True


# ─────────────────────────────────────────────────────────────────────────────
# Inline formatting parser
# ─────────────────────────────────────────────────────────────────────────────

def apply_inline(para, text: str, base_bold=False, base_italic=False,
                 base_size=None, base_color=None, font_name=None):
    """
    Parse inline markdown (bold, italic, strikethrough, inline code, links)
    and add runs to the paragraph.
    """
    # Pattern order matters — more specific first
    pattern = re.compile(
        r"(\*\*\*|___)"           # bold+italic (***  or ___)
        r"|(\*\*|__)"             # bold
        r"|(\*|_)"                # italic
        r"|(~~)"                  # strikethrough
        r"|(`.+?`)"               # inline code
        r"|(\[([^\]]*)\]\(([^)]*)\))"  # link [text](url)
        r"|(<([^>]+)>)"           # autolink <url>
        r"|(\\\S)"                # escaped char
    )

    pos = 0
    bold = base_bold
    italic = base_italic
    strike = False
    toggle_stack = []  # track open toggles

    def emit_run(txt):
        if not txt:
            return
        run = para.add_run(txt)
        run.bold = bold
        run.italic = italic
        run.font.strike = strike
        if base_size:
            run.font.size = base_size
        if base_color:
            run.font.color.rgb = base_color
        if font_name:
            run.font.name = font_name

    i = 0
    segments = []
    for m in pattern.finditer(text):
        if m.start() > i:
            segments.append(("text", text[i:m.start()]))
        if m.group(1):   # bold+italic
            segments.append(("toggle_bi", m.group(1)))
        elif m.group(2): # bold
            segments.append(("toggle_b", m.group(2)))
        elif m.group(3): # italic
            segments.append(("toggle_i", m.group(3)))
        elif m.group(4): # strikethrough
            segments.append(("toggle_s", m.group(4)))
        elif m.group(5): # inline code
            code_text = m.group(5)[1:-1]  # strip backticks
            segments.append(("code", code_text))
        elif m.group(6): # link
            segments.append(("link", m.group(7), m.group(8)))
        elif m.group(9): # autolink
            url = m.group(10)
            segments.append(("link", url, url))
        elif m.group(11): # escaped
            segments.append(("text", m.group(11)[1:]))
        i = m.end()

    if i < len(text):
        segments.append(("text", text[i:]))

    # Render segments with toggle state
    b, it, s = base_bold, base_italic, False
    for seg in segments:
        kind = seg[0]
        if kind == "text":
            if not seg[1]:
                continue
            run = para.add_run(seg[1])
            run.bold = b
            run.italic = it
            run.font.strike = s
            if base_size:
                run.font.size = base_size
            if base_color:
                run.font.color.rgb = base_color
            if font_name:
                run.font.name = font_name
        elif kind == "toggle_bi":
            b = not b
            it = not it
        elif kind == "toggle_b":
            b = not b
        elif kind == "toggle_i":
            it = not it
        elif kind == "toggle_s":
            s = not s
        elif kind == "code":
            run = para.add_run(seg[1])
            run.font.name = "Courier New"
            run.font.size = Pt(10)
            # light gray highlight for inline code
            rPr = run._r.get_or_add_rPr()
            highlight = OxmlElement("w:highlight")
            highlight.set(qn("w:val"), "lightGray")
            rPr.append(highlight)
        elif kind == "link":
            link_text, url = seg[1], seg[2]
            try:
                add_hyperlink(para, link_text or url, url)
            except Exception:
                run = para.add_run(link_text or url)
                run.font.color.rgb = RGBColor(0x00, 0x56, 0xB2)


# ─────────────────────────────────────────────────────────────────────────────
# AST → DOCX renderer
# ─────────────────────────────────────────────────────────────────────────────

class DocxRenderer(mistune.BaseRenderer):
    """Renders a mistune AST directly into a python-docx Document."""

    def __init__(self, doc: Document, md_path: Path):
        super().__init__()
        self.doc = doc
        self.md_path = md_path  # for resolving relative image paths
        self._list_level = 0
        self._list_type_stack = []  # "bullet" or "number"
        self._in_table_header = False

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _inline_tokens_to_text(self, tokens) -> str:
        """Flatten inline tokens to plain text (for tables / headings fallback)."""
        parts = []
        for tok in tokens:
            if isinstance(tok, dict):
                t = tok.get("type", "")
                if t == "text":
                    parts.append(tok.get("raw", ""))
                elif t in ("strong", "em", "codespan", "del"):
                    parts.append(self._inline_tokens_to_text(tok.get("children", [])))
                elif t == "softline":
                    parts.append(" ")
                elif t == "linebreak":
                    parts.append("\n")
                elif t == "link":
                    parts.append(self._inline_tokens_to_text(tok.get("children", [])))
                elif t == "image":
                    parts.append(tok.get("attrs", {}).get("alt", ""))
                elif t == "raw":
                    # strip HTML tags
                    raw = tok.get("raw", "")
                    parts.append(re.sub(r"<[^>]+>", "", raw))
            elif isinstance(tok, str):
                parts.append(tok)
        return "".join(parts)

    def _render_inline_to_para(self, para, tokens, **kwargs):
        """Render a list of inline tokens to runs in a paragraph."""
        for tok in tokens:
            if isinstance(tok, dict):
                t = tok.get("type", "")
                if t == "text":
                    apply_inline(para, tok.get("raw", ""), **kwargs)
                elif t == "softline":
                    para.add_run(" ")
                elif t == "linebreak":
                    para.add_run().add_break()
                elif t == "strong":
                    self._render_inline_to_para(para, tok.get("children", []),
                                                 base_bold=True, **{k: v for k, v in kwargs.items() if k != "base_bold"})
                elif t == "em":
                    self._render_inline_to_para(para, tok.get("children", []),
                                                 base_italic=True, **{k: v for k, v in kwargs.items() if k != "base_italic"})
                elif t == "del":
                    child_tokens = tok.get("children", [])
                    for ct in child_tokens:
                        if isinstance(ct, dict) and ct.get("type") == "text":
                            run = para.add_run(ct.get("raw", ""))
                            run.font.strike = True
                elif t == "codespan":
                    run = para.add_run(tok.get("raw", ""))
                    run.font.name = "Courier New"
                    run.font.size = Pt(10)
                    rPr = run._r.get_or_add_rPr()
                    highlight = OxmlElement("w:highlight")
                    highlight.set(qn("w:val"), "lightGray")
                    rPr.append(highlight)
                elif t == "link":
                    url = tok.get("attrs", {}).get("url", "#")
                    link_text = self._inline_tokens_to_text(tok.get("children", []))
                    try:
                        add_hyperlink(para, link_text or url, url)
                    except Exception:
                        run = para.add_run(link_text or url)
                        run.font.color.rgb = RGBColor(0x00, 0x56, 0xB2)
                elif t == "image":
                    attrs = tok.get("attrs", {})
                    src = attrs.get("url", "")
                    alt = attrs.get("alt", "")
                    self._add_image_to_para(para, src, alt)
                elif t == "raw":
                    raw = tok.get("raw", "")
                    clean = re.sub(r"<[^>]+>", "", raw)
                    if clean.strip():
                        para.add_run(clean)
            elif isinstance(tok, str):
                apply_inline(para, tok, **kwargs)

    def _add_image_to_para(self, para, src: str, alt: str = ""):
        """Try to add an image; fall back to alt text if not found."""
        # Resolve relative path
        img_path = Path(src)
        if not img_path.is_absolute():
            img_path = self.md_path.parent / src

        if img_path.exists():
            try:
                run = para.add_run()
                run.add_picture(str(img_path), width=Inches(5.5))
                return
            except Exception:
                pass

        # URL or missing: show alt text as italic
        display = f"[Imagem: {alt or src}]"
        run = para.add_run(display)
        run.italic = True
        run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

    # ── Block elements ────────────────────────────────────────────────────────

    def heading(self, token, state):
        level = token.get("attrs", {}).get("level", 1)
        children = token.get("children", [])
        style_name = f"Heading {min(level, 6)}"
        para = self.doc.add_paragraph(style=style_name)
        self._render_inline_to_para(para, children)
        return ""

    def paragraph(self, token, state):
        children = token.get("children", [])
        para = self.doc.add_paragraph(style="Normal")
        self._render_inline_to_para(para, children)
        return ""

    def blank_line(self, token, state):
        return ""

    def thematic_break(self, token, state):
        add_horizontal_rule(self.doc)
        return ""

    def block_code(self, token, state):
        """Fenced or indented code block."""
        raw = token.get("raw", "")
        # Split into lines preserving empty lines
        lines = raw.split("\n")
        # Remove trailing empty line if present
        if lines and lines[-1] == "":
            lines = lines[:-1]

        for i, line in enumerate(lines):
            para = self.doc.add_paragraph(style="Code Block")
            set_paragraph_shading(para, "F2F2F2")
            para.paragraph_format.space_before = Pt(0) if i > 0 else Pt(6)
            para.paragraph_format.space_after = Pt(0) if i < len(lines) - 1 else Pt(6)
            run = para.add_run(line)
            run.font.name = "Courier New"
            run.font.size = Pt(9.5)
        return ""

    def block_quote(self, token, state):
        children = token.get("children", [])
        for child in children:
            if isinstance(child, dict):
                ct = child.get("type", "")
                if ct == "paragraph":
                    para = self.doc.add_paragraph(style="Block Quote")
                    set_left_border(para)
                    self._render_inline_to_para(para, child.get("children", []))
                else:
                    # Nested block quote or other element
                    self._render_token(child, state)
        return ""

    def list(self, token, state):
        ordered = token.get("attrs", {}).get("ordered", False)
        self._list_type_stack.append("number" if ordered else "bullet")
        self._list_level += 1
        children = token.get("children", [])
        for child in children:
            self._render_token(child, state)
        self._list_level -= 1
        self._list_type_stack.pop()
        return ""

    def list_item(self, token, state):
        list_type = self._list_type_stack[-1] if self._list_type_stack else "bullet"
        level = self._list_level  # 1-based

        # Choose style based on level and type
        if list_type == "bullet":
            style = "List Bullet" if level == 1 else f"List Bullet {min(level, 3)}"
        else:
            style = "List Number" if level == 1 else f"List Number {min(level, 3)}"

        children = token.get("children", [])
        para = None

        for child in children:
            if isinstance(child, dict):
                ct = child.get("type", "")
                if ct == "block_text":
                    para = self.doc.add_paragraph(style=style)
                    self._render_inline_to_para(para, child.get("children", []))
                elif ct == "paragraph":
                    if para is None:
                        para = self.doc.add_paragraph(style=style)
                        self._render_inline_to_para(para, child.get("children", []))
                    else:
                        # continuation paragraph
                        p2 = self.doc.add_paragraph(style="Normal")
                        p2.paragraph_format.left_indent = Cm(level * 0.63)
                        self._render_inline_to_para(p2, child.get("children", []))
                elif ct == "list":
                    # nested list
                    self._render_token(child, state)
                else:
                    self._render_token(child, state)
            elif isinstance(child, str) and child.strip():
                if para is None:
                    para = self.doc.add_paragraph(style=style)
                    apply_inline(para, child)
        return ""

    def table(self, token, state):
        children = token.get("children", [])
        if not children:
            return ""

        # Separate head and body
        head_rows = []
        body_rows = []
        for child in children:
            if isinstance(child, dict):
                if child.get("type") == "table_head":
                    for row in child.get("children", []):
                        head_rows.append(row)
                elif child.get("type") == "table_body":
                    for row in child.get("children", []):
                        body_rows.append(row)

        if not head_rows and not body_rows:
            return ""

        # Count columns from header
        num_cols = max(
            len(r.get("children", [])) for r in (head_rows + body_rows)
            if isinstance(r, dict)
        ) if (head_rows or body_rows) else 1

        num_rows = len(head_rows) + len(body_rows)
        table = self.doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        set_table_style(table)

        row_idx = 0
        for row_tok in head_rows:
            if isinstance(row_tok, dict):
                self._fill_table_row(table.rows[row_idx], row_tok, is_header=True)
                row_idx += 1

        for row_tok in body_rows:
            if isinstance(row_tok, dict):
                self._fill_table_row(table.rows[row_idx], row_tok, is_header=False)
                row_idx += 1

        if head_rows:
            shade_table_header(table.rows[0])

        # Space after table
        self.doc.add_paragraph(style="Normal").paragraph_format.space_after = Pt(4)
        return ""

    def _fill_table_row(self, row, row_token, is_header: bool):
        cells = row_token.get("children", [])
        for col_idx, cell_tok in enumerate(cells):
            if col_idx >= len(row.cells):
                break
            cell = row.cells[col_idx]
            # Clear default empty paragraph
            for p in cell.paragraphs:
                p._element.getparent().remove(p._element)

            para = cell.add_paragraph()
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.space_before = Pt(2)

            if isinstance(cell_tok, dict):
                align = cell_tok.get("attrs", {}).get("align")
                if align == "center":
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif align == "right":
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT

                inline_children = cell_tok.get("children", [])
                self._render_inline_to_para(para, inline_children,
                                             base_bold=is_header)

    def _render_token(self, token, state):
        """Dispatch a single token to its renderer method."""
        if not isinstance(token, dict):
            return
        t = token.get("type", "")
        method = getattr(self, t, None)
        if method:
            method(token, state)
        elif t in ("block_text",):
            para = self.doc.add_paragraph(style="Normal")
            self._render_inline_to_para(para, token.get("children", []))

    # ── Raw HTML (strip tags) ─────────────────────────────────────────────────
    def block_html(self, token, state):
        raw = token.get("raw", "")
        # Handle <br>, <hr>
        if re.search(r"<hr\s*/?>", raw, re.I):
            add_horizontal_rule(self.doc)
            return ""
        # Strip other HTML
        clean = re.sub(r"<[^>]+>", "", raw).strip()
        if clean:
            para = self.doc.add_paragraph(style="Normal")
            para.add_run(clean)
        return ""

    def inline_html(self, token, state):
        return ""

    # ── Finalize ──────────────────────────────────────────────────────────────
    def finalize(self, data, state):
        return ""


# ─────────────────────────────────────────────────────────────────────────────
# Core conversion function
# ─────────────────────────────────────────────────────────────────────────────

def convert_md_to_docx(md_path: Path, docx_path: Path):
    """Convert a single Markdown file to DOCX."""
    md_text = md_path.read_text(encoding="utf-8", errors="replace")

    # Strip YAML front matter
    md_text = re.sub(r"^---\s*\n.*?\n---\s*\n", "", md_text, flags=re.DOTALL)

    doc = Document()

    # Page margins (Word default: 2.54cm = 1 inch)
    section = doc.sections[0]
    section.page_width = Cm(21.0)    # A4
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)

    setup_styles(doc)

    renderer = DocxRenderer(doc=doc, md_path=md_path)
    md_parser = mistune.create_markdown(
        renderer=renderer,
        plugins=["strikethrough", "table", "footnotes", "task_lists"],
    )
    md_parser(md_text)

    # Remove leading empty paragraph that Word always creates
    if doc.paragraphs and doc.paragraphs[0].text == "":
        p = doc.paragraphs[0]._element
        p.getparent().remove(p)

    doc.save(str(docx_path))


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="MD to DOCX Converter by Jair Lima",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  md2docx                        # Convert all .md in current folder
  md2docx README.md              # Convert specific file
  md2docx README.md --force      # Overwrite existing DOCX
  md2docx --folder /path/to/dir  # Convert all .md in another folder
        """,
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="Specific .md file to convert (optional)",
    )
    parser.add_argument(
        "--folder",
        default=".",
        help="Folder to scan for .md files (default: current directory)",
    )
    parser.add_argument(
        "--force", "-f",
        action="store_true",
        help="Overwrite existing DOCX files",
    )
    parser.add_argument(
        "--output", "-o",
        help="Output folder for DOCX files (default: same as source)",
    )
    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None

    # ── Single file mode ──────────────────────────────────────────────────────
    if args.file:
        md_path = Path(args.file).resolve()
        if not md_path.exists():
            print(f"[ERRO] Arquivo não encontrado: {md_path}")
            sys.exit(1)
        if md_path.suffix.lower() != ".md":
            print(f"[ERRO] O arquivo deve ter extensão .md: {md_path}")
            sys.exit(1)

        out_dir = output_dir or md_path.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        docx_path = out_dir / (md_path.stem + ".docx")

        if docx_path.exists() and not args.force:
            print(f"[PULA]  {md_path.name} → já existe {docx_path.name}")
            sys.exit(0)

        print(f"[CONV]  {md_path.name} → {docx_path.name}")
        convert_md_to_docx(md_path, docx_path)
        print(f"[OK]    Salvo em: {docx_path}")
        sys.exit(0)

    # ── Folder scan mode ──────────────────────────────────────────────────────
    folder = Path(args.folder).resolve()
    if not folder.is_dir():
        print(f"[ERRO] Pasta não encontrada: {folder}")
        sys.exit(1)

    md_files = sorted(folder.glob("*.md"))
    if not md_files:
        print(f"[INFO] Nenhum arquivo .md encontrado em: {folder}")
        sys.exit(0)

    converted = 0
    skipped = 0
    errors = 0

    print(f"[SCAN]  {folder}")
    print(f"[INFO]  {len(md_files)} arquivo(s) .md encontrado(s)\n")

    for md_path in md_files:
        out_dir = output_dir or md_path.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        docx_path = out_dir / (md_path.stem + ".docx")

        if docx_path.exists() and not args.force:
            print(f"  [PULA]  {md_path.name}")
            skipped += 1
            continue

        try:
            print(f"  [CONV]  {md_path.name} → {docx_path.name}")
            convert_md_to_docx(md_path, docx_path)
            print(f"  [OK]    Salvo.")
            converted += 1
        except Exception as e:
            print(f"  [ERRO]  {md_path.name}: {e}")
            errors += 1

    print(f"\n[FIM]   Convertidos: {converted}  |  Pulados: {skipped}  |  Erros: {errors}")


if __name__ == "__main__":
    main()
