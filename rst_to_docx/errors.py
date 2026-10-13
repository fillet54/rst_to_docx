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
        details = "\n  ".join(self.messages)
        super().__init__(f"RST source has {len(self.messages)} error(s):\n  {details}")


class UnsupportedNodeError(RstToDocxError):
    """The document uses an RST construct the renderer cannot output."""

    def __init__(self, node, context=None):
        self.node = node
        msg = f"unsupported RST element <{node.tagname}>{_location(node)}"
        super().__init__(f"{msg} ({context})" if context else msg)


class TemplateError(RstToDocxError):
    """The Word template is missing a marker or is otherwise unusable."""


class UndefinedVariableError(RstToDocxError):
    """Template ``{{placeholders}}`` referenced variables that were not given."""

    def __init__(self, names):
        self.names = list(names)
        super().__init__(f"template references undefined variable(s): {', '.join(self.names)}")


class MissingStyleWarning(UserWarning):
    """A Word style was not found in the template and a fallback was used."""


class RstWarning(UserWarning):
    """docutils reported a warning-level problem in the RST source."""


def _location(node):
    """`` at <source>:<line>`` for the node or its nearest ancestor with a line."""
    while node is not None:
        if getattr(node, "line", None) is not None:
            return f" at {getattr(node, 'source', None) or '<string>'}:{node.line}"
        node = getattr(node, "parent", None)
    return ""
