"""Exceptions and warnings raised by rst_to_docx."""


class RstToDocxError(Exception):
    """Base class for all rst_to_docx errors."""


class RstError(RstToDocxError):
    """The RST source contained errors (bad markup, undefined variables, ...).

    ``messages`` holds one ``"<source>:<line>: <text>"`` string per problem so
    callers can report every issue at once rather than fixing them one by one.
    """

    def __init__(self, messages):
        self.messages = list(messages)
        super().__init__(
            "RST source has %d error(s):\n  %s"
            % (len(self.messages), "\n  ".join(self.messages))
        )


class UnsupportedNodeError(RstToDocxError):
    """The document uses an RST construct the renderer cannot output."""

    def __init__(self, node, context=None):
        self.node = node
        where = _location(node)
        msg = "unsupported RST element <%s>%s" % (node.tagname, where)
        if context:
            msg += " (%s)" % context
        super().__init__(msg)


class TemplateError(RstToDocxError):
    """The Word template is missing a marker or is otherwise unusable."""


class UndefinedVariableError(RstToDocxError):
    """Template ``{{placeholders}}`` referenced variables that were not given."""

    def __init__(self, names):
        self.names = list(names)
        super().__init__(
            "template references undefined variable(s): %s" % ", ".join(self.names)
        )


class MissingStyleWarning(UserWarning):
    """A Word style was not found in the template and a fallback was used."""


class RstWarning(UserWarning):
    """docutils reported a warning-level problem in the RST source."""


def _location(node):
    line = None
    source = None
    while node is not None and line is None:
        line = getattr(node, "line", None)
        source = source or getattr(node, "source", None)
        node = getattr(node, "parent", None)
    if line is None:
        return ""
    return " at %s:%s" % (source or "<string>", line)
