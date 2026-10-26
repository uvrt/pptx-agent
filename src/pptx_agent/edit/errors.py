"""The library's refusal of a value it was given: :class:`ArgumentError`.

Every check the editing API makes of what it is passed -- a colour, a length, a spec's
keys, a slide that has no such placeholder -- raises an :class:`ArgumentError`, whose
message says what is wrong and what would do.  It is a :class:`ValueError`, as these
refusals always were, so ``except ValueError`` still catches it.  A check of a value's
type raises :class:`ArgumentTypeError`, both an :class:`ArgumentError` and the
:class:`TypeError` it always was.

Both are of ooxml-edit's family (:mod:`ooxml_edit.errors`): an ``ArgumentError`` here is
an ``ooxml_edit.errors.ArgumentError``, an ``ArgumentTypeError`` an
``ooxml_edit.errors.ArgumentTypeError``, so a refusal is one class whichever library
raised it, in a deck or in a document.

The tool layer tells the model ``invalid_arguments`` for these
(:func:`pptx_agent.tools.common.library_errors`), as the dispatcher does for ooxml-edit's
own refusals (the last node of a diagram, a cell reference that is not one); ooxml-edit's
``UnitError`` (a value out of its range) is ``unit``.  A ``ValueError`` or ``TypeError``
Python raises itself is not a refusal and goes to the dispatcher's own mapping, as in the
other formats' tools.
"""

from __future__ import annotations

from ooxml_edit.errors import ArgumentError as _ArgumentError
from ooxml_edit.errors import ArgumentTypeError as _ArgumentTypeError


class ArgumentError(_ArgumentError):
    """A value the library refuses; the message says what is wrong and what would do."""


class ArgumentTypeError(ArgumentError, _ArgumentTypeError):
    """A value of a type the library does not take there; the message says which it does."""


__all__ = ["ArgumentError", "ArgumentTypeError"]
