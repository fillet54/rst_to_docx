"""List numbering support (numbering.xml).

Bullet lists rely on the numbering already attached to the template's
"List Bullet N" styles. Numbered lists must restart at each new RST list, so
each one gets a fresh ``w:num`` that points at the style's abstract numbering
with a ``startOverride``. When a style carries no numbering (or the RST asks
for letters/roman numerals) a generated abstract definition is used instead.
"""

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

ENUM_FORMATS = {
    "arabic": ("decimal", "%{n}."),
    "loweralpha": ("lowerLetter", "%{n}."),
    "upperalpha": ("upperLetter", "%{n}."),
    "lowerroman": ("lowerRoman", "%{n}."),
    "upperroman": ("upperRoman", "%{n}."),
}
BULLET_CHARS = ["•", "o", "▪"]


def style_numbering(style):
    """Return ``(numId, ilvl)`` from a paragraph style or its ancestors, else None."""
    while style is not None:
        pPr = style.element.pPr
        if pPr is not None and pPr.numPr is not None and pPr.numPr.numId is not None:
            ilvl = pPr.numPr.ilvl.val if pPr.numPr.ilvl is not None else 0
            num_id = pPr.numPr.numId.val
            return (num_id, ilvl) if num_id else None
        style = style.base_style
    return None


def set_paragraph_numbering(paragraph, num_id, ilvl):
    numPr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    numPr.get_or_add_ilvl().val = ilvl
    numPr.get_or_add_numId().val = num_id


class Numbering:
    def __init__(self, document):
        self._el = document.part.numbering_part.element
        self._generated = {}

    def restart(self, num_id, ilvl, start=1):
        """New ``w:num`` sharing ``num_id``'s abstract definition, restarted."""
        num = self._el.num_having_numId(num_id)
        new = self._el.add_num(num.abstractNumId.val)
        override = new.add_lvlOverride(ilvl=ilvl)
        override.add_startOverride(start)
        return new.numId

    def generated(self, kind, ilvl=0, start=1):
        """Number id for a generated list definition; levels map to nesting depth.

        ``kind`` is ``"bullet"`` or an RST enumtype such as ``"loweralpha"``.
        Bullets share one ``w:num``; enumerated lists get a fresh restarted one.
        """
        abstract_id = self._generated.get(kind)
        if abstract_id is None:
            abstract_id = self._add_abstract(kind)
            self._generated[kind] = abstract_id
            if kind == "bullet":
                self._generated["bullet-num"] = self._el.add_num(abstract_id).numId
        if kind == "bullet":
            return self._generated["bullet-num"]
        num = self._el.add_num(abstract_id)
        for lvl in range(9):
            num.add_lvlOverride(ilvl=lvl).add_startOverride(start if lvl == ilvl else 1)
        return num.numId

    def _add_abstract(self, kind):
        existing = [int(a.get(qn("w:abstractNumId"))) for a in self._el.findall(qn("w:abstractNum"))]
        abstract_id = max(existing, default=-1) + 1
        levels = []
        for lvl in range(9):
            indent = 720 * (lvl + 1)
            if kind == "bullet":
                fmt, text = "bullet", BULLET_CHARS[lvl % len(BULLET_CHARS)]
            else:
                fmt, pattern = ENUM_FORMATS.get(kind, ENUM_FORMATS["arabic"])
                text = pattern.replace("{n}", str(lvl + 1))
            levels.append(
                '<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/>'
                '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
                '<w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr></w:lvl>'
                % (lvl, fmt, text, indent)
            )
        abstract = parse_xml(
            '<w:abstractNum %s w:abstractNumId="%d"><w:multiLevelType w:val="hybridMultilevel"/>'
            "%s</w:abstractNum>" % (nsdecls("w"), abstract_id, "".join(levels))
        )
        # Schema order: every abstractNum precedes the first num.
        first_num = self._el.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract)
        else:
            self._el.append(abstract)
        return abstract_id
