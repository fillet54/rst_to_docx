"""Render reStructuredText into a Word (.docx) template with python-docx.

    from rst_to_docx import render
    render(Path("svd.rst"), "template.docx", "out.docx",
           variables={"release": {"version": "1.4.2"}},
           jira_url="https://jira.example.com/browse/")

Custom roles and directives are plain docutils ones (register them with
``docutils.parsers.rst.roles`` / ``directives``); if they produce new node
types, add a renderer for them with :func:`register_handler`.
"""

from . import extensions as _extensions
from .api import parse, render
from .errors import (
    MissingStyleWarning, RstError, RstToDocxError, RstWarning, TemplateError,
    UndefinedVariableError, UnsupportedNodeError,
)
from .renderer import Renderer, register_handler

_extensions.register()

__all__ = [
    "render", "parse", "register_handler", "Renderer",
    "RstToDocxError", "RstError", "UnsupportedNodeError", "TemplateError",
    "UndefinedVariableError", "MissingStyleWarning", "RstWarning",
]
