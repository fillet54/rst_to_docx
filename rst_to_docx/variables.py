"""Variable lookup shared by the :var: role, var-table and template placeholders."""

from collections.abc import Mapping, Sequence

_MISSING = object()


class VariableLookupError(KeyError):
    def __init__(self, path, reason):
        super().__init__(path)
        self.path = path
        self.reason = reason

    def __str__(self):
        return f"variable {self.path!r} {self.reason}"


class NotScalarError(VariableLookupError):
    """The variable exists but is a mapping or list where text was expected."""


def lookup(variables, path):
    """Resolve a dotted ``path`` such as ``release.version`` or ``files.0.name``.

    Each segment indexes a mapping by key, a sequence by integer index, or
    falls back to attribute access (so dataclasses and namedtuples work).
    """
    value = variables
    parts = path.split(".")
    for i, part in enumerate(parts):
        if isinstance(value, Mapping):
            value = value.get(part, _MISSING)
        elif isinstance(value, Sequence) and not isinstance(value, str) and part.isdigit():
            value = value[int(part)] if int(part) < len(value) else _MISSING
        elif not part.startswith("_"):
            value = getattr(value, part, _MISSING)
        else:
            value = _MISSING
        if value is _MISSING:
            where = f" (failed at {'.'.join(parts[:i + 1])!r})" if i else ""
            raise VariableLookupError(path, f"is not defined{where}")
    return value


def format_scalar(path, value):
    """Return ``value`` as display text, rejecting structured values."""
    if isinstance(value, (Mapping, list, tuple, set, frozenset)):
        raise NotScalarError(path, f"is a {type(value).__name__}, not a scalar; use var-table for lists")
    return "" if value is None else str(value)


def lookup_scalar(variables, path):
    return format_scalar(path, lookup(variables, path))
