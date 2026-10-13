"""Variable lookup shared by the :var: role, var-table and template placeholders."""

from collections.abc import Mapping, Sequence

_MISSING = object()


class VariableLookupError(KeyError):
    def __init__(self, path, reason):
        self.path = path
        self.reason = reason
        super().__init__(path)

    def __str__(self):
        return "variable %r %s" % (self.path, self.reason)


def lookup(variables, path):
    """Resolve a dotted ``path`` such as ``release.version`` or ``files.0.name``.

    Each segment indexes a mapping by key, a sequence by integer index, or
    falls back to attribute access (so dataclasses and namedtuples work).
    """
    value = variables
    walked = []
    for part in path.split("."):
        walked.append(part)
        nxt = _MISSING
        if isinstance(value, Mapping):
            nxt = value.get(part, _MISSING)
        elif isinstance(value, Sequence) and not isinstance(value, str) and part.isdigit():
            idx = int(part)
            if idx < len(value):
                nxt = value[idx]
        elif not part.startswith("_"):
            nxt = getattr(value, part, _MISSING)
        if nxt is _MISSING:
            where = "" if len(walked) == 1 else " (failed at %r)" % ".".join(walked)
            raise VariableLookupError(path, "is not defined" + where)
        value = nxt
    return value


def is_scalar(value):
    return not isinstance(value, (Mapping, list, tuple, set, frozenset))


def format_scalar(path, value):
    """Return ``value`` as display text, rejecting structured values."""
    if not is_scalar(value):
        raise VariableLookupError(
            path, "is a %s, not a scalar; use var-table for lists" % type(value).__name__
        )
    return "" if value is None else str(value)


def lookup_scalar(variables, path):
    return format_scalar(path, lookup(variables, path))
