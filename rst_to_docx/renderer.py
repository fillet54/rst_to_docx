"""Walks a docutils doctree and writes python-docx content before an anchor.

Dispatch is by node class through a handler registry (see :func:`register_handler`),
so callers can add renderers for their own docutils nodes. A handler is
``handler(renderer, node)`` and uses the renderer's helpers:

* ``renderer.new_paragraph(style_chain)`` - add a paragraph in the current container
* ``renderer.add_text(text)`` - add a run with the current inline formatting
* ``renderer.formatting(**flags)`` - context manager for inline formatting
* ``renderer.render_children(node)`` - recurse
"""

import hashlib
import re
from contextlib import contextmanager

from docutils import nodes
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.run import Run

from . import styles as S
from .errors import UnsupportedNodeError
from .extensions import ref_node
from .numbering import Numbering, set_paragraph_numbering, style_numbering

_handlers = {}

MONO_FONT = "Courier New"
ADMONITIONS = (
    nodes.attention, nodes.caution, nodes.danger, nodes.error, nodes.hint,
    nodes.important, nodes.note, nodes.tip, nodes.warning,
)


def register_handler(node_class, handler=None):
    """Register ``handler(renderer, node)`` for ``node_class`` (usable as a decorator)."""
    if handler is None:
        def decorator(fn):
            _handlers[node_class] = fn
            return fn
        return decorator
    _handlers[node_class] = handler
    return handler


def handler_for(node):
    for cls in type(node).__mro__:
        if cls in _handlers:
            return _handlers[cls]
    return None


# -- containers ------------------------------------------------------------------


class AnchorContainer:
    """Inserts block content immediately before the template's {{content}} paragraph."""

    def __init__(self, document, anchor):
        self.document = document
        self.anchor = anchor

    def add_paragraph(self):
        return self.anchor.insert_paragraph_before()

    def add_table(self, rows, cols):
        # Word merges directly adjacent tables into one; keep them apart.
        prev = self.anchor._p.getprevious()
        if prev is not None and prev.tag == qn("w:tbl"):
            self.add_paragraph()
        table = self.document.add_table(rows, cols)
        self.anchor._p.addprevious(table._tbl)
        return table

    @property
    def width(self):
        section = self.document.sections[-1]
        return section.page_width - section.left_margin - section.right_margin


class CellContainer:
    """Block content inside a table cell (reuses the cell's initial empty paragraph)."""

    def __init__(self, cell):
        self.cell = cell
        self._reusable = cell.paragraphs[0] if not cell.paragraphs[0].text else None

    def add_paragraph(self):
        if self._reusable is not None:
            p, self._reusable = self._reusable, None
            return p
        return self.cell.add_paragraph()

    def add_table(self, rows, cols):
        if self._reusable is not None:
            # python-docx appends tables after existing content; drop the unused
            # leading paragraph so the cell doesn't start with a blank line.
            p = self._reusable._p
            self._reusable = None
            table = self.cell.add_table(rows, cols)
            p.getparent().remove(p)
        else:
            last = self.cell._tc[-1]
            if last.tag == qn("w:tbl"):
                self.cell.add_paragraph()
            table = self.cell.add_table(rows, cols)
        # Cells must end with a paragraph; python-docx added one, reuse it next.
        self._reusable = self.cell.paragraphs[-1]
        return table

    @property
    def width(self):
        return self.cell.width


# -- renderer --------------------------------------------------------------------


class _ListState:
    def __init__(self, kind, depth, num_id=None, ilvl=0):
        self.kind = kind
        self.depth = depth
        self.num_id = num_id
        self.ilvl = ilvl
        self.first_paragraph = False


class Renderer:
    def __init__(self, document, anchor, doctree):
        self.document = document
        self.doctree = doctree
        self.styles = S.StyleResolver(document)
        self.numbering = Numbering(document)
        self.containers = [AnchorContainer(document, anchor)]
        self.contexts = []          # "cell" or _ListState, innermost last
        self.paragraph = None       # current python-docx Paragraph (inline mode)
        self.run_parent = None      # element runs are appended to (w:p or w:hyperlink)
        self.fmt = {}
        self.section_depth = 0
        self.pending_ids = []
        self._bookmark_id = self._max_bookmark_id() + 1
        self._bookmarks = {}

    # -- public helpers ------------------------------------------------------

    @property
    def container(self):
        return self.containers[-1]

    def render(self, node):
        if isinstance(node, nodes.Element) and node.get("ids"):
            self.pending_ids.extend(node["ids"])
            if self.paragraph is not None:
                self._emit_bookmarks()
        handler = handler_for(node)
        if handler is None:
            raise UnsupportedNodeError(node)
        handler(self, node)

    def render_children(self, node):
        for child in node.children:
            self.render(child)

    def new_paragraph(self, style_chain=None):
        para = self.container.add_paragraph()
        if style_chain:
            style = self.styles.get(style_chain)
            if style is not None:
                para.style = style
        self.paragraph = para
        self.run_parent = para._p
        self._emit_bookmarks()
        return para

    @contextmanager
    def inline_paragraph(self, style_chain=None):
        """Open a paragraph, render inline content into it, then close it."""
        para = self.new_paragraph(style_chain)
        try:
            yield para
        finally:
            self.paragraph = None
            self.run_parent = None

    @contextmanager
    def formatting(self, **flags):
        saved = self.fmt
        self.fmt = dict(saved, **flags)
        try:
            yield
        finally:
            self.fmt = saved

    @contextmanager
    def in_container(self, container, context="cell"):
        self.containers.append(container)
        self.contexts.append(context)
        saved = (self.paragraph, self.run_parent)
        self.paragraph = self.run_parent = None
        try:
            yield container
        finally:
            self.containers.pop()
            self.contexts.pop()
            self.paragraph, self.run_parent = saved

    def add_text(self, text):
        if self.paragraph is None:
            raise UnsupportedNodeError(nodes.Text(text), "text outside a paragraph")
        r = OxmlElement("w:r")
        self.run_parent.append(r)
        run = Run(r, self.paragraph)
        run.text = text
        f = self.fmt
        if f.get("bold"):
            run.bold = True
        if f.get("italic"):
            run.italic = True
        if f.get("underline"):
            run.underline = True
        if f.get("strike"):
            run.font.strike = True
        if f.get("superscript"):
            run.font.superscript = True
        if f.get("subscript"):
            run.font.subscript = True
        if f.get("code"):
            style = self.styles.get("HTML Code", S.CHARACTER)
            if style is not None:
                run.style = style
            else:
                run.font.name = MONO_FONT
        if f.get("hyperlink"):
            style = self.styles.get("Hyperlink", S.CHARACTER)
            if style is not None and not f.get("code"):
                run.style = style
            else:
                run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                run.underline = True
        return run

    # -- bookmarks / hyperlinks ----------------------------------------------

    def bookmark_name(self, node_id):
        name = self._bookmarks.get(node_id)
        if name is None:
            name = "_r2d_" + re.sub(r"[^A-Za-z0-9_]", "_", node_id)
            if len(name) > 40:
                digest = hashlib.sha1(node_id.encode()).hexdigest()[:8]
                name = name[:31] + "_" + digest
            self._bookmarks[node_id] = name
        return name

    def _emit_bookmarks(self):
        if not self.pending_ids or self.paragraph is None:
            return
        for node_id in self.pending_ids:
            start = OxmlElement("w:bookmarkStart")
            start.set(qn("w:id"), str(self._bookmark_id))
            start.set(qn("w:name"), self.bookmark_name(node_id))
            end = OxmlElement("w:bookmarkEnd")
            end.set(qn("w:id"), str(self._bookmark_id))
            self._bookmark_id += 1
            self.run_parent.append(start)
            self.run_parent.append(end)
        self.pending_ids = []

    def _max_bookmark_id(self):
        ids = [int(b.get(qn("w:id"), 0)) for b in self.document.element.body.iter(qn("w:bookmarkStart"))]
        return max(ids, default=0)

    @contextmanager
    def hyperlink(self, url=None, anchor_id=None):
        link = OxmlElement("w:hyperlink")
        if url is not None:
            rid = self.paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
            link.set(qn("r:id"), rid)
        else:
            link.set(qn("w:anchor"), self.bookmark_name(anchor_id))
        self.run_parent.append(link)
        saved = self.run_parent
        self.run_parent = link
        try:
            with self.formatting(hyperlink=True):
                yield link
        finally:
            self.run_parent = saved

    # -- paragraph style for context -----------------------------------------

    def body_paragraph_style(self):
        ctx = self.contexts[-1] if self.contexts else None
        if isinstance(ctx, _ListState):
            if ctx.first_paragraph:
                return None  # handled by list item
            return S.list_continue_chain(ctx.depth)
        return None

    def list_depth(self):
        return sum(1 for c in self.contexts if isinstance(c, _ListState))


# -- block handlers --------------------------------------------------------------


def _skip(renderer, node):
    pass


for _cls in (nodes.comment, nodes.target, nodes.substitution_definition,
             nodes.system_message, nodes.pending):
    register_handler(_cls, _skip)


@register_handler(nodes.document)
@register_handler(nodes.compound)
@register_handler(nodes.container)
def _children(renderer, node):
    renderer.render_children(node)


@register_handler(nodes.section)
def _section(renderer, node):
    renderer.section_depth += 1
    try:
        renderer.render_children(node)
    finally:
        renderer.section_depth -= 1


@register_handler(nodes.title)
def _title(renderer, node):
    parent = node.parent
    if isinstance(parent, nodes.section):
        chain = S.heading_chain(max(renderer.section_depth, 1))
    elif isinstance(parent, nodes.table):
        chain = ["Caption", "Normal"]
    else:
        raise UnsupportedNodeError(node, "title inside <%s>" % parent.tagname)
    with renderer.inline_paragraph(chain):
        renderer.render_children(node)


@register_handler(nodes.paragraph)
def _paragraph(renderer, node):
    ctx = renderer.contexts[-1] if renderer.contexts else None
    if isinstance(ctx, _ListState) and ctx.first_paragraph:
        ctx.first_paragraph = False
        para_ctx = renderer.inline_paragraph(S.list_chain(ctx.kind, ctx.depth))
        with para_ctx as para:
            if ctx.num_id is not None:
                set_paragraph_numbering(para, ctx.num_id, ctx.ilvl)
            elif style_numbering(para.style) is None:
                # The resolved style (maybe a fallback) has no numbering.
                kind = "bullet" if ctx.kind == "bullet" else "arabic"
                ctx.num_id, ctx.ilvl = renderer.numbering.generated(kind, ctx.depth - 1), ctx.depth - 1
                set_paragraph_numbering(para, ctx.num_id, ctx.ilvl)
            renderer.render_children(node)
        return
    with renderer.inline_paragraph(renderer.body_paragraph_style()):
        renderer.render_children(node)


def _list(renderer, node, kind):
    depth = renderer.list_depth() + 1
    state = _ListState(kind, depth)
    if kind == "enum":
        enumtype = node.get("enumtype", "arabic")
        start = node.get("start", 1)
        style = renderer.styles.get(S.list_chain("enum", depth))
        numbering = style_numbering(style) if style is not None else None
        if enumtype == "arabic" and numbering is not None:
            num_id, ilvl = numbering
            state.num_id, state.ilvl = renderer.numbering.restart(num_id, ilvl, start), ilvl
        else:
            state.ilvl = depth - 1
            state.num_id = renderer.numbering.generated(enumtype, state.ilvl, start)
    renderer.contexts.append(state)
    try:
        for item in node.children:
            if not isinstance(item, nodes.list_item):
                raise UnsupportedNodeError(item, "inside a list")
            renderer.render(item)
    finally:
        renderer.contexts.pop()


register_handler(nodes.bullet_list, lambda r, n: _list(r, n, "bullet"))
register_handler(nodes.enumerated_list, lambda r, n: _list(r, n, "enum"))


@register_handler(nodes.list_item)
def _list_item(renderer, node):
    state = renderer.contexts[-1]
    state.first_paragraph = True
    if not node.children or not isinstance(node.children[0], nodes.paragraph):
        # Empty item or one starting with a non-paragraph: still emit the bullet.
        node.insert(0, nodes.paragraph())
    renderer.render_children(node)
    state.first_paragraph = False


@register_handler(nodes.literal_block)
@register_handler(nodes.doctest_block)
def _literal_block(renderer, node):
    style = renderer.styles.get(list(S.MONOSPACE_PARAGRAPH_STYLES) + ["Normal"])
    with renderer.inline_paragraph() as para:
        if style is not None:
            para.style = style
        run = renderer.add_text(node.astext())
        if style is None or style.name not in S.MONOSPACE_PARAGRAPH_STYLES:
            run.font.name = MONO_FONT
            run.font.size = Pt(9)


@register_handler(nodes.block_quote)
def _block_quote(renderer, node):
    # Indented text in RST; usually accidental inside lists. Render as Quote.
    for child in node.children:
        if isinstance(child, nodes.paragraph):
            with renderer.inline_paragraph(["Quote", "Normal"]):
                renderer.render_children(child)
        elif isinstance(child, nodes.attribution):
            with renderer.inline_paragraph(["Quote", "Normal"]):
                renderer.add_text("— ")
                renderer.render_children(child)
        else:
            renderer.render(child)


def _admonition(renderer, node):
    if isinstance(node, nodes.admonition):
        title_node = node.next_node(nodes.title)
        body = [c for c in node.children if c is not title_node]
    else:
        title_node = None
        body = list(node.children)
    table = renderer.container.add_table(1, 1)
    style = renderer.styles.get("Table Grid", S.TABLE)
    if style is not None:
        table.style = style
    with renderer.in_container(CellContainer(table.cell(0, 0))):
        with renderer.inline_paragraph():
            with renderer.formatting(bold=True):
                if title_node is not None:
                    renderer.render_children(title_node)
                else:
                    renderer.add_text(node.tagname.capitalize())
        for child in body:
            renderer.render(child)


register_handler(nodes.admonition, _admonition)
for _cls in ADMONITIONS:
    register_handler(_cls, _admonition)


# -- tables ------------------------------------------------------------------------


def _mark_header_row(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


@register_handler(nodes.table)
def _table(renderer, node):
    title = None
    tgroup = None
    for child in node.children:
        if isinstance(child, nodes.title):
            title = child
        elif isinstance(child, nodes.tgroup):
            tgroup = child
        else:
            raise UnsupportedNodeError(child, "inside a table")
    if tgroup is None:
        raise UnsupportedNodeError(node, "table without tgroup")

    ncols = tgroup["cols"]
    colwidths = []
    head_rows, body_rows = [], []
    for child in tgroup.children:
        if isinstance(child, nodes.colspec):
            colwidths.append(child.get("colwidth", 1))
        elif isinstance(child, nodes.thead):
            head_rows.extend(child.children)
        elif isinstance(child, nodes.tbody):
            body_rows.extend(child.children)
    rows = head_rows + body_rows
    if not rows:
        raise UnsupportedNodeError(node, "table with no rows")

    # Lay entries out on a grid, honouring row/column spans.
    occupied = set()
    placements = []
    for r, row in enumerate(rows):
        c = 0
        for entry in row.children:
            while (r, c) in occupied:
                c += 1
            mr, mc = entry.get("morerows", 0), entry.get("morecols", 0)
            placements.append((r, c, mr, mc, entry))
            for rr in range(r, r + mr + 1):
                for cc in range(c, c + mc + 1):
                    occupied.add((rr, cc))
            c += mc + 1

    if title is not None:
        renderer.render(title)

    table = renderer.container.add_table(len(rows), ncols)
    style = renderer.styles.get("Table Grid", S.TABLE)
    if style is not None:
        table.style = style

    for r, c, mr, mc, _ in placements:
        if mr or mc:
            table.cell(r, c).merge(table.cell(r + mr, c + mc))

    if len(colwidths) == ncols and isinstance(renderer.container, AnchorContainer):
        table.autofit = False
        total = float(sum(colwidths))
        widths = [int(renderer.container.width * w / total) for w in colwidths]
        for i, w in enumerate(widths):
            table.columns[i].width = w
        for r, c, mr, mc, _ in placements:
            table.cell(r, c).width = sum(widths[c:c + mc + 1])

    for i in range(len(head_rows)):
        _mark_header_row(table.rows[i])

    for r, c, mr, mc, entry in placements:
        with renderer.in_container(CellContainer(table.cell(r, c))):
            if r < len(head_rows):
                with renderer.formatting(bold=True):
                    renderer.render_children(entry)
            else:
                renderer.render_children(entry)


# -- inline handlers -------------------------------------------------------------


@register_handler(nodes.Text)
def _text(renderer, node):
    text = node.astext().replace("\n", " ")
    if text:
        renderer.add_text(text)


def _formatted(**flags):
    def handler(renderer, node):
        with renderer.formatting(**flags):
            renderer.render_children(node)
    return handler


register_handler(nodes.strong, _formatted(bold=True))
register_handler(nodes.emphasis, _formatted(italic=True))
register_handler(nodes.title_reference, _formatted(italic=True))
register_handler(nodes.literal, _formatted(code=True))
register_handler(nodes.superscript, _formatted(superscript=True))
register_handler(nodes.subscript, _formatted(subscript=True))

INLINE_CLASS_FLAGS = {"underline": "underline", "strike": "strike"}


@register_handler(nodes.inline)
def _inline(renderer, node):
    flags = {INLINE_CLASS_FLAGS[c]: True for c in node.get("classes", []) if c in INLINE_CLASS_FLAGS}
    with renderer.formatting(**flags):
        renderer.render_children(node)


@register_handler(nodes.reference)
def _reference(renderer, node):
    if "refuri" in node:
        with renderer.hyperlink(url=node["refuri"]):
            renderer.render_children(node)
    elif "refid" in node:
        with renderer.hyperlink(anchor_id=node["refid"]):
            renderer.render_children(node)
    else:
        raise UnsupportedNodeError(node, "unresolved reference")


@register_handler(ref_node)
def _ref(renderer, node):
    # Resolved by api.resolve_refs before rendering; refid is always set here.
    with renderer.hyperlink(anchor_id=node["refid"]):
        renderer.render_children(node)
