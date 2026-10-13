"""Resolve built-in Word style names against the template, with fallbacks.

Word only writes a built-in style into styles.xml once it has been used, so a
template may lack e.g. "Heading 5" or "List Bullet 3". Rather than fail, each
lookup walks a fallback chain and emits a single MissingStyleWarning per
missing name.
"""

import warnings

from docx.enum.style import WD_STYLE_TYPE

from .errors import MissingStyleWarning

PARAGRAPH = WD_STYLE_TYPE.PARAGRAPH
CHARACTER = WD_STYLE_TYPE.CHARACTER
TABLE = WD_STYLE_TYPE.TABLE

MONOSPACE_PARAGRAPH_STYLES = ("Plain Text", "Macro Text", "HTML Preformatted")


def heading_chain(level):
    return [f"Heading {n}" for n in range(max(level, 1), 0, -1)] + ["Normal"]


def appendix_chain(level):
    """Appendix N styles, falling back to the ordinary headings."""
    return [f"Appendix {n}" for n in range(max(level, 1), 0, -1)] + heading_chain(level)


def _numbered_chain(base, level, final):
    return [f"{base} {n}" for n in range(level, 1, -1)] + [base, *final]


def list_chain(kind, level):
    base = {"bullet": "List Bullet", "enum": "List Number"}[kind]
    return _numbered_chain(base, level, ["List Paragraph", "Normal"])


def list_continue_chain(level):
    return _numbered_chain("List Continue", level, ["Normal"])


class StyleResolver:
    def __init__(self, document):
        self.document = document
        self._cache = {}
        self._warned = set()

    def get(self, chain, style_type=PARAGRAPH):
        """Return the first style in ``chain`` present in the template.

        ``chain`` is a name or a list of names. Returns ``None`` when nothing in
        the chain exists (callers then rely on direct formatting or defaults).
        """
        if isinstance(chain, str):
            chain = [chain]
        key = (tuple(chain), style_type)
        if key in self._cache:
            return self._cache[key]
        found = next(
            (style for name in chain if (style := self._lookup(name, style_type)) is not None),
            None,
        )
        wanted = chain[0]
        if (found is None or found.name != wanted) and wanted not in self._warned:
            self._warned.add(wanted)
            fallback = repr(found.name) if found is not None else "direct formatting"
            warnings.warn(
                f"style {wanted!r} not found in template; using {fallback}. Use the style "
                f"once on the template's instructions page so Word saves it.",
                MissingStyleWarning,
                stacklevel=4,
            )
        self._cache[key] = found
        return found

    def _lookup(self, name, style_type):
        try:
            style = self.document.styles[name]
        except KeyError:
            return None
        return style if style.type == style_type else None
