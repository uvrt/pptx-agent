"""The library's refusal of a value it was given: :class:`ArgumentError`.

Every check the editing API makes of what it is passed -- a colour, a length, a spec's
keys, a slide that has no such placeholder -- raises an :class:`ArgumentError`, whose
message says what is wrong and what would do.  It is a :class:`ValueError`, as these
refusals always were, so ``except ValueError`` still catches it.  A check of a value's
type raises :class:`ArgumentTypeError`, both an :class:`ArgumentError` and the
:class:`TypeError` it always was.

The tool layer tells the model ``invalid_arguments`` for these and for nothing else
(:func:`pptx_agent.tools.common.library_errors`): a ``ValueError`` or ``TypeError``
Python raises itself, or a sibling library's, is not the library's refusal and goes to
the dispatcher's own mapping, as in the other formats' tools.
"""

from __future__ import annotations


class ArgumentError(ValueError):
    """A value the library refuses; the message says what is wrong and what would do."""


class ArgumentTypeError(ArgumentError, TypeError):
    """A value of a type the library does not take there; the message says which it does."""


__all__ = ["ArgumentError", "ArgumentTypeError"]
