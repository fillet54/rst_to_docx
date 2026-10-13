"""Custom docutils roles and directives.

Everything here runs at *parse* time and produces ordinary docutils nodes
(Text, reference, table, ...) wherever possible, so the renderer needs no
special knowledge of them. Render-time options (variables, Jira URL) are read
from ``document.settings``, which :func:`rst_to_docx.render` populates.

Roles
    :var:`release.version`     value from the variables dict
    :jira:`PROJ-123`           hyperlink to the Jira issue
    :jira:`text <PROJ-123>`    same, with custom link text
    :ref:`label`               link to an explicit target / section
    :ref:`text <label>`        same, with custom link text
    :underline:`text`, :strike:`text`

Directives
    .. var-table:: files       table built from a list of dicts
"""

import re
from collections.abc import Iterable

from docutils import nodes
from docutils.parsers.rst import Directive, directives, roles

from .variables import VariableLookupError, format_scalar, lookup, lookup_scalar

SETTING_VARIABLES = "rst2docx_variables"
SETTING_JIRA_URL = "rst2docx_jira_url"
CELL_FORMATS = {"jira", "code"}

# docutils escapes backslashes as \x00, so "\<" is not an explicit title.
_EXPLICIT_TITLE = re.compile(r"^(.+?)\s*(?<!\x00)<(.*?)>$", re.DOTALL)
_JIRA_KEY = re.compile(r"^[A-Za-z][A-Za-z0-9_]*-\d+$")


class ref_node(nodes.Inline, nodes.TextElement):
    """Unresolved :ref: cross-reference; resolved in api._resolve_refs."""


def _split_explicit_title(text):
    """``"text <target>"`` -> ``("text", "target")``; ``"target"`` -> ``(None, "target")``."""
    if m := _EXPLICIT_TITLE.match(text):
        return m.group(1), m.group(2).strip()
    return None, text.strip()


def _setting(document, name):
    return getattr(document.settings, name, None)


def _role_error(inliner, rawtext, lineno, message):
    msg = inliner.reporter.error(message, line=lineno)
    return [inliner.problematic(rawtext, rawtext, msg)], [msg]


def jira_reference(base_url, key, text=None):
    """Reference node linking ``key`` to ``base_url`` (which may contain ``{key}``)."""
    if not _JIRA_KEY.match(key):
        raise ValueError(f"{key!r} is not a Jira issue key")
    url = base_url.replace("{key}", key) if "{key}" in base_url else f"{base_url.rstrip('/')}/{key}"
    return nodes.reference(text or key, text or key, refuri=url)


# -- roles -----------------------------------------------------------------------


def var_role(name, rawtext, text, lineno, inliner, options=None, content=None):
    variables = _setting(inliner.document, SETTING_VARIABLES) or {}
    try:
        value = lookup_scalar(variables, text.strip())
    except VariableLookupError as exc:
        return _role_error(inliner, rawtext, lineno, f":var: {exc}")
    return [nodes.inline(rawtext, value, classes=["var"])], []


def jira_role(name, rawtext, text, lineno, inliner, options=None, content=None):
    base_url = _setting(inliner.document, SETTING_JIRA_URL)
    if not base_url:
        return _role_error(inliner, rawtext, lineno, ":jira: used but no jira_url was passed to render()")
    title, key = _split_explicit_title(text)
    try:
        return [jira_reference(base_url, key, title)], []
    except ValueError as exc:
        return _role_error(inliner, rawtext, lineno, f":jira: {exc}")


def ref_role(name, rawtext, text, lineno, inliner, options=None, content=None):
    title, label = _split_explicit_title(text)
    node = ref_node(rawtext, title or "", reflabel=label, explicit_title=title is not None)
    node.line = lineno
    node.source = inliner.document.current_source
    return [node], []


def _class_role(cls):
    def role(name, rawtext, text, lineno, inliner, options=None, content=None):
        return [nodes.inline(rawtext, text, classes=[cls])], []
    return role


# -- var-table -------------------------------------------------------------------


def parse_columns(spec):
    """Parse ``key=Header[:fmt], ...`` into ``[(key, header, fmt), ...]``.

    Entries are separated by commas or newlines. A trailing ``:jira`` or
    ``:code`` on the header selects a cell format; any other colon is kept as
    part of the header text.
    """
    columns = []
    for raw in filter(None, (s.strip() for s in re.split(r"[,\n]", spec))):
        key, _, header = (s.strip() for s in raw.partition("="))
        if not key:
            raise ValueError(f"empty column key in {raw!r}")
        header = header or key
        fmt = None
        head, _, suffix = header.rpartition(":")
        if head and suffix.strip() in CELL_FORMATS:
            header, fmt = head.strip(), suffix.strip()
        columns.append((key, header, fmt))
    if not columns:
        raise ValueError("no columns given")
    return columns


def _positive_int_list(argument):
    values = [int(v) for v in re.split(r"[,\s]+", argument.strip()) if v]
    if any(v <= 0 for v in values):
        raise ValueError("widths must be positive integers")
    return values


def _row(paragraphs, **entry_attrs):
    return nodes.row("", *(nodes.entry("", para, **entry_attrs) for para in paragraphs))


class VarTable(Directive):
    """Render a list of mappings from the variables dict as a table.

    .. var-table:: changelog
       :columns: key=Ticket:jira, summary=Summary, author=Author
       :widths: 15 65 20
       :title: Changes in this release
       :empty: No changes.
    """

    required_arguments = 1
    option_spec = {
        "columns": directives.unchanged_required,
        "widths": _positive_int_list,
        "title": directives.unchanged,
        "empty": directives.unchanged,
        "class": directives.class_option,
    }

    def fail(self, message):
        return self.error(f"var-table {self.arguments[0]!r}: {message}")

    def run(self):
        path = self.arguments[0].strip()
        if "columns" not in self.options:
            raise self.fail("the :columns: option is required")
        try:
            columns = parse_columns(self.options["columns"])
        except ValueError as exc:
            raise self.fail(f"bad :columns: ({exc})")
        widths = self.options.get("widths", [1] * len(columns))
        if len(widths) != len(columns):
            raise self.fail(f"{len(widths)} widths given for {len(columns)} columns")

        document = self.state.document
        jira_url = _setting(document, SETTING_JIRA_URL)
        if not jira_url and any(fmt == "jira" for _, _, fmt in columns):
            raise self.fail("a :jira column is used but no jira_url was passed to render()")
        try:
            records = lookup(_setting(document, SETTING_VARIABLES) or {}, path)
        except VariableLookupError as exc:
            raise self.fail(str(exc))
        if records is None:
            records = []
        if isinstance(records, (str, bytes)) or not isinstance(records, Iterable):
            raise self.fail("variable is not a list")

        tbody = nodes.tbody()
        for index, record in enumerate(records):
            try:
                tbody += _row(self._cell(record, key, fmt, jira_url) for key, _, fmt in columns)
            except (VariableLookupError, ValueError) as exc:
                raise self.fail(f"row {index}: {exc}")
        if not tbody.children and "empty" in self.options:
            text = self.options["empty"]
            tbody += _row([nodes.paragraph(text, text)], morecols=len(columns) - 1)

        table = nodes.table(classes=self.options.get("class", []))
        messages = []
        if title := self.options.get("title"):
            title_nodes, messages = self.state.inline_text(title, self.lineno)
            table += nodes.title(title, "", *title_nodes)
        table += nodes.tgroup(
            "",
            *(nodes.colspec(colwidth=w) for w in widths),
            nodes.thead("", _row(nodes.paragraph(h, h) for _, h, _ in columns)),
            tbody,
            cols=len(columns),
        )
        self.add_name(table)
        table.line = self.lineno
        return [table, *messages]

    @staticmethod
    def _cell(record, key, fmt, jira_url):
        value = lookup(record, key)
        para = nodes.paragraph()
        for i, item in enumerate(value if isinstance(value, (list, tuple)) else [value]):
            if i:
                para += nodes.Text(", ")
            text = format_scalar(key, item)
            if fmt == "jira":
                para += jira_reference(jira_url, text)
            elif fmt == "code":
                para += nodes.literal(text, text)
            else:
                para += nodes.Text(text)
        return para


def register():
    roles.register_local_role("var", var_role)
    roles.register_local_role("jira", jira_role)
    roles.register_local_role("ref", ref_role)
    roles.register_local_role("underline", _class_role("underline"))
    roles.register_local_role("strike", _class_role("strike"))
    directives.register_directive("var-table", VarTable)
