"""List numbering support (numbering.xml).

Bullet lists rely on the numbering already attached to the template's
"List Bullet N" styles. Numbered lists must restart at each new RST list, so
each one gets a fresh ``w:num`` that points at the style's abstract numbering
with a ``startOverride``. When a style carries no numbering (or the RST asks
for letters/roman numerals) a generated abstract definition is used instead.
"""

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.parts.numbering import NumberingPart

NUMBER_FORMATS = {
    "arabic": "decimal",
    "loweralpha": "lowerLetter",
    "upperalpha": "upperLetter",
    "lowerroman": "lowerRoman",
    "upperroman": "upperRoman",
}
BULLET_CHARS = ("•", "o", "▪")


def set_paragraph_numbering(paragraph, num_id, ilvl):
    numPr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    numPr.get_or_add_ilvl().val = ilvl
    numPr.get_or_add_numId().val = num_id


def _numbering_element(document):
    """The ``w:numbering`` element, adding numbering.xml if the template has none.

    python-docx can't create the part itself (``NumberingPart.new`` raises
    NotImplementedError), and Word omits it from documents that never used a list.
    """
    try:
        return document.part.numbering_part.element
    except NotImplementedError:
        part = NumberingPart(
            PackURI("/word/numbering.xml"), CT.WML_NUMBERING,
            parse_xml(f"<w:numbering {nsdecls('w')}/>"), document.part.package,
        )
        document.part.relate_to(part, RT.NUMBERING)
        return part.element


class Numbering:
    def __init__(self, document):
        self._el = _numbering_element(document)
        self._abstract_ids = {}
        self._bullet_num_id = None

    def for_style(self, style):
        """``(numId, ilvl)`` from a paragraph style or its ancestors, else None.

        Returns None when the referenced ``w:num`` doesn't exist, since Word
        would then show no number at all.
        """
        while style is not None:
            pPr = style.element.pPr
            numPr = pPr.numPr if pPr is not None else None
            if numPr is not None and numPr.numId is not None:
                num_id = numPr.numId.val
                if not num_id or self._num(num_id) is None:
                    return None  # numId 0 explicitly switches numbering off
                return num_id, numPr.ilvl.val if numPr.ilvl is not None else 0
            style = style.base_style
        return None

    def restart(self, num_id, ilvl, start=1):
        """New ``w:num`` sharing ``num_id``'s abstract definition, restarted at ``start``."""
        new = self._el.add_num(self._num(num_id).abstractNumId.val)
        new.add_lvlOverride(ilvl=ilvl).add_startOverride(start)
        return new.numId

    def generated(self, kind, ilvl=0, start=1):
        """Number id for a generated list definition; levels map to nesting depth.

        ``kind`` is ``"bullet"`` or an RST enumtype such as ``"loweralpha"``.
        Bullets share one ``w:num``; enumerated lists get a fresh restarted one.
        """
        if kind not in self._abstract_ids:
            self._abstract_ids[kind] = self._add_abstract(kind)
        abstract_id = self._abstract_ids[kind]
        if kind == "bullet":
            if self._bullet_num_id is None:
                self._bullet_num_id = self._el.add_num(abstract_id).numId
            return self._bullet_num_id
        num = self._el.add_num(abstract_id)
        num.add_lvlOverride(ilvl=ilvl).add_startOverride(start)
        return num.numId

    def _num(self, num_id):
        return next((n for n in self._el.num_lst if n.numId == num_id), None)

    def _add_abstract(self, kind):
        abstract_id = 1 + max(
            (int(a.get(qn("w:abstractNumId"))) for a in self._el.iterchildren(qn("w:abstractNum"))),
            default=-1,
        )
        levels = []
        for lvl in range(9):
            if kind == "bullet":
                fmt, text = "bullet", BULLET_CHARS[lvl % len(BULLET_CHARS)]
            else:
                fmt, text = NUMBER_FORMATS.get(kind, "decimal"), f"%{lvl + 1}."
            levels.append(
                f'<w:lvl w:ilvl="{lvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{720 * (lvl + 1)}" w:hanging="360"/></w:pPr></w:lvl>'
            )
        abstract = parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
            f'<w:multiLevelType w:val="hybridMultilevel"/>{"".join(levels)}</w:abstractNum>'
        )
        # Schema order: every abstractNum precedes the first num.
        first_num = self._el.find(qn("w:num"))
        if first_num is not None:
            first_num.addprevious(abstract)
        else:
            self._el.append(abstract)
        return abstract_id
