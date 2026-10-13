"""Operations on the Word template: instructions page, markers, placeholders."""

import re

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .errors import TemplateError, UndefinedVariableError
from .variables import VariableLookupError, lookup_scalar

CONTENT_MARKER = "{{content}}"
END_INSTRUCTIONS_MARKER = "{{end-instructions}}"
RESERVED = {"content", "end-instructions"}

PLACEHOLDER = re.compile(r"\{\{\s*([A-Za-z_][\w.\-]*)\s*\}\}")

W_P = qn("w:p")
W_T = qn("w:t")


def paragraph_text(p):
    """Text of a ``w:p`` element, excluding paragraphs nested in text boxes."""
    return "".join(t.text or "" for t in _own_text_elements(p))


def _own_text_elements(p):
    for t in p.iter(W_T):
        parent = t.getparent()
        while parent is not None and parent.tag != W_P:
            parent = parent.getparent()
        if parent is p:
            yield t


def _is_page_break_only(el):
    if el.tag != W_P or paragraph_text(el).strip():
        return False
    return any(br.get(qn("w:type")) == "page" for br in el.iter(qn("w:br")))


def find_end_instructions(body):
    """Index of the top-level body child holding the end-instructions marker."""
    for i, child in enumerate(body.iterchildren()):
        if child.tag == W_P and paragraph_text(child).strip() == END_INSTRUCTIONS_MARKER:
            return i
    return None


def instructions_elements(body):
    """Top-level body elements that make up the instructions page (may be empty)."""
    idx = find_end_instructions(body)
    if idx is None:
        return []
    children = list(body.iterchildren())
    end = idx + 1
    if end < len(children) and _is_page_break_only(children[end]):
        end += 1
    return children[:end]


def strip_instructions(document):
    body = document.element.body
    elements = instructions_elements(body)
    if not elements:
        raise TemplateError(
            "strip_instructions=True but the template has no paragraph containing only %r; add the "
            "marker at the end of the instructions page or pass strip_instructions=False"
            % END_INSTRUCTIONS_MARKER
        )
    for el in elements:
        body.remove(el)


def find_content_anchor(document, skip=()):
    """The ``{{content}}`` paragraph (as a python-docx Paragraph)."""
    from docx.text.paragraph import Paragraph

    skip_ids = {id(el) for el in skip}
    body = document.element.body
    matches = []
    for child in body.iterchildren():
        if id(child) in skip_ids:
            continue
        for p in child.iter(W_P):
            if paragraph_text(p).strip() == CONTENT_MARKER:
                matches.append(p)
    if not matches:
        raise TemplateError("template has no paragraph containing only %r" % CONTENT_MARKER)
    if len(matches) > 1:
        raise TemplateError("template has %d %r paragraphs; expected one" % (len(matches), CONTENT_MARKER))
    return Paragraph(matches[0], document._body)


def remove_anchor(paragraph):
    p = paragraph._p
    pPr = p.pPr
    if pPr is not None and pPr.sectPr is not None:
        # The marker paragraph ends a section; keep it (and the section break)
        # but drop its text.
        for child in list(p):
            if child is not pPr:
                p.remove(child)
    else:
        p.getparent().remove(p)


# -- placeholders ----------------------------------------------------------------


def _template_roots(document, skip):
    body = document.element.body
    skip_ids = {id(el) for el in skip}
    for child in body.iterchildren():
        if id(child) not in skip_ids:
            yield child
    seen = set()
    for rel in document.part.rels.values():
        if rel.is_external or rel.reltype not in (RT.HEADER, RT.FOOTER):
            continue
        part = rel.target_part
        if id(part) not in seen:
            seen.add(id(part))
            yield part.element


def substitute_placeholders(document, variables, skip=()):
    """Replace ``{{dotted.name}}`` in body, headers and footers.

    Word often splits text like ``{{version}}`` over several runs, so matching
    is done on the paragraph's combined text and the replacement is written into
    the run where the placeholder starts (keeping that run's formatting).
    """
    undefined = []
    for root in _template_roots(document, skip):
        for p in list(root.iter(W_P)):
            _substitute_paragraph(p, variables, undefined)
    if undefined:
        raise UndefinedVariableError(sorted(set(undefined)))


def _substitute_paragraph(p, variables, undefined):
    texts = list(_own_text_elements(p))
    full = "".join(t.text or "" for t in texts)
    if "{{" not in full:
        return
    matches = [m for m in PLACEHOLDER.finditer(full) if m.group(1) not in RESERVED]
    if not matches:
        return
    # Map each character offset to (text element, offset within it).
    spans = []
    pos = 0
    for t in texts:
        n = len(t.text or "")
        spans.append((pos, pos + n, t))
        pos += n
    for m in reversed(matches):
        name = m.group(1)
        try:
            value = lookup_scalar(variables, name)
        except VariableLookupError as exc:
            undefined.append("%s (%s)" % (name, exc.reason) if "scalar" in exc.reason else name)
            continue
        _replace_range(spans, m.start(), m.end(), value)


def _replace_range(spans, start, end, value):
    first = True
    for s, e, t in spans:
        if e <= start or s >= end:
            continue
        text = t.text or ""
        a = max(start, s) - s
        b = min(end, e) - s
        if first:
            t.text = text[:a] + value + text[b:]
            first = False
        else:
            t.text = text[:a] + text[b:]
        t.set(qn("xml:space"), "preserve")


# -- settings --------------------------------------------------------------------

# Elements that must follow w:updateFields in CT_Settings (schema order).
_AFTER_UPDATE_FIELDS = [
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
    "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
]


def has_fields(document):
    roots = list(_template_roots(document, ()))
    tags = (qn("w:fldChar"), qn("w:fldSimple"))
    return any(el.tag in tags for root in roots for el in root.iter(*tags))


def set_update_fields(document):
    settings = document.settings.element
    existing = settings.find(qn("w:updateFields"))
    if existing is None:
        existing = OxmlElement("w:updateFields")
        successors = {qn(tag) for tag in _AFTER_UPDATE_FIELDS}
        for child in settings.iterchildren():
            if child.tag in successors:
                child.addprevious(existing)
                break
        else:
            settings.append(existing)
    existing.set(qn("w:val"), "true")
