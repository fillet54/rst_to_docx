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

from docutils import nodes
from docutils.parsers.rst import Directive, directives, roles

from .variables import VariableLookupError, format_scalar, lookup, lookup_scalar

SETTING_VARIABLES = "rst2docx_variables"
SETTING_JIRA_URL = "rst2docx_jira_url"

_EXPLICIT_TITLE = re.compile(r"^(.+?)\s*(?<!\x00)<(.*?)>$", re.DOTALL)
_JIRA_KEY = re.compile(r"^[A-Za-z][A-Za-z0-9_]*-\d+$")


class ref_node(nodes.Inline, nodes.TextElement):
    """Unresolved :ref: cross-reference; resolved by the renderer."""


def _split_explicit_title(text):
    m = _EXPLICIT_TITLE.match(text)
    if m:
        return True, m.group(1), m.group(2)
    return False, text, text


def _settings(inliner_or_state, name):
    return getattr(inliner_or_state.document.settings, name, None)


def _role_error(inliner, rawtext, lineno, message):
    msg = inliner.reporter.error(message, line=lineno)
    prb = inliner.problematic(rawtext, rawtext, msg)
    return [prb], [msg]


def jira_url_for(base, key):
    if "{key}" in base:
        return base.replace("{key}", key)
    return base.rstrip("/") + "/" + key


def make_jira_reference(base, key, text=None):
    return nodes.reference(text or key, text or key, refuri=jira_url_for(base, key))


# -- roles -------------------------------------------------------------------


def var_role(name, rawtext, text, lineno, inliner, options=None, content=None):
    variables = _settings(inliner, SETTING_VARIABLES) or {}
    path = text.strip()
    try:
        value = lookup_scalar(variables, path)
    except VariableLookupError as exc:
        return _role_error(inliner, rawtext, lineno, ":var: %s" % exc)
    return [nodes.inline(rawtext, value, classes=["var"])], []


def jira_role(name, rawtext, text, lineno, inliner, options=None, content=None):
    base = _settings(inliner, SETTING_JIRA_URL)
    if not base:
        return _role_error(
            inliner, rawtext, lineno, ":jira: used but no jira_url was passed to render()"
        )
    has_title, title, key = _split_explicit_title(text)
    key = key.strip()
    if not _JIRA_KEY.match(key):
        return _role_error(inliner, rawtext, lineno, ":jira: %r is not a Jira issue key" % key)
    return [make_jira_reference(base, key, title if has_title else None)], []


def ref_role(name, rawtext, text, lineno, inliner, options=None, content=None):
    has_title, title, label = _split_explicit_title(text)
    node = ref_node(rawtext, title if has_title else "", reflabel=label.strip(),
                    explicit_title=has_title)
    node.line = lineno
    node.source = inliner.document.current_source
    return [node], []


def _class_role(cls):
    def role(name, rawtext, text, lineno, inliner, options=None, content=None):
        return [nodes.inline(rawtext, text, classes=[cls])], []
    return role


# -- var-table -----------------------------------------------------------------


FORMATTERS = ("jira", "code")


def parse_columns(spec):
    """Parse ``key=Header[:fmt], ...`` into ``[(key, header, fmt), ...]``.

    Entries are separated by commas or newlines. A trailing ``:jira`` or
    ``:code`` on the header selects a cell formatter; any other colon is kept
    as part of the header text.
    """
    columns = []
    for raw in re.split(r"[,\n]", spec):
        raw = raw.strip()
        if not raw:
            continue
        key, sep, header = raw.partition("=")
        key = key.strip()
        header = header.strip() if sep else key
        fmt = None
        head, colon, suffix = header.rpartition(":")
        if colon and suffix.strip() in FORMATTERS:
            header, fmt = head.strip(), suffix.strip()
        if not key:
            raise ValueError("empty column key in %r" % raw)
        columns.append((key, header, fmt))
    if not columns:
        raise ValueError("no columns given")
    return columns


def _positive_int_list(argument):
    values = [int(v) for v in re.split(r"[,\s]+", argument.strip()) if v]
    if any(v <= 0 for v in values):
        raise ValueError("widths must be positive integers")
    return values


class VarTable(Directive):
    """Render a list of mappings from the variables dict as a table.

    .. var-table:: changelog
       :columns: key=Ticket:jira, summary=Summary, author=Author
       :widths: 15 65 20
       :title: Changes in this release
       :empty: No changes.
    """

    required_arguments = 1
    optional_arguments = 0
    has_content = False
    option_spec = {
        "columns": directives.unchanged_required,
        "widths": _positive_int_list,
        "title": directives.unchanged,
        "empty": directives.unchanged,
        "class": directives.class_option,
    }

    def run(self):
        path = self.arguments[0].strip()
        if "columns" not in self.options:
            raise self.error("var-table %r: the :columns: option is required" % path)
        try:
            columns = parse_columns(self.options["columns"])
        except ValueError as exc:
            raise self.error("var-table %r: bad :columns: (%s)" % (path, exc))
        widths = self.options.get("widths") or [1] * len(columns)
        if len(widths) != len(columns):
            raise self.error(
                "var-table %r: %d widths given for %d columns" % (path, len(widths), len(columns))
            )

        variables = getattr(self.state.document.settings, SETTING_VARIABLES, None) or {}
        jira_base = getattr(self.state.document.settings, SETTING_JIRA_URL, None)
        try:
            rows = lookup(variables, path)
        except VariableLookupError as exc:
            raise self.error("var-table: %s" % exc)
        if rows is None:
            rows = []
        if isinstance(rows, (str, bytes)) or not hasattr(rows, "__iter__"):
            raise self.error("var-table: variable %r is not a list" % path)

        if any(fmt == "jira" for _, _, fmt in columns) and not jira_base:
            raise self.error(
                "var-table %r: a :jira column is used but no jira_url was passed to render()"
                % path
            )

        messages = []
        table = nodes.table(classes=self.options.get("class", []))
        if self.options.get("title"):
            title_text = self.options["title"]
            title_nodes, messages = self.state.inline_text(title_text, self.lineno)
            table += nodes.title(title_text, "", *title_nodes)
        tgroup = nodes.tgroup(cols=len(columns))
        table += tgroup
        for w in widths:
            tgroup += nodes.colspec(colwidth=w)

        thead = nodes.thead()
        tgroup += thead
        thead += self._row([nodes.paragraph(h, h) for _, h, _ in columns])

        tbody = nodes.tbody()
        tgroup += tbody
        count = 0
        for index, record in enumerate(rows):
            count += 1
            cells = []
            for key, _, fmt in columns:
                try:
                    value = lookup(record, key)
                except VariableLookupError as exc:
                    raise self.error("var-table %r row %d: %s" % (path, index, exc))
                try:
                    cells.append(self._cell(key, value, fmt, jira_base))
                except ValueError as exc:
                    raise self.error("var-table %r row %d: %s" % (path, index, exc))
            tbody += self._row(cells)

        if count == 0 and self.options.get("empty"):
            text = self.options["empty"]
            entry = nodes.entry(morecols=len(columns) - 1)
            entry += nodes.paragraph(text, text)
            row = nodes.row()
            row += entry
            tbody += row

        self.add_name(table)
        table.line = self.lineno
        return [table] + messages

    @staticmethod
    def _row(paragraphs):
        row = nodes.row()
        for para in paragraphs:
            entry = nodes.entry()
            entry += para
            row += entry
        return row

    @staticmethod
    def _cell(key, value, fmt, jira_base):
        values = list(value) if isinstance(value, (list, tuple)) else [value]
        para = nodes.paragraph()
        for i, item in enumerate(values):
            if i:
                para += nodes.Text(", ")
            text = format_scalar(key, item)
            if fmt == "jira":
                if not _JIRA_KEY.match(text):
                    raise ValueError("%r is not a Jira issue key" % text)
                para += make_jira_reference(jira_base, text)
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
