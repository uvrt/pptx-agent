"""Driving PowerPoint, and recovering when it will not be driven.

PowerPoint is the only authoritative answer to "will this file open?".  Getting that answer
reliably means handling the case where the answer is *no*, because a damaged file raises an
app-modal repair dialog ("Herstellen" / "Repair") and PowerPoint stops servicing AppleEvents
while it is up.

That has two consequences worth stating, because both cost an afternoon to rediscover:

1. **The export script cannot dismiss the dialog.**  It is blocked inside `open`, so it never
   regains control.  Anything that clears the dialog has to run in a different process.
2. **The dialog outlives the failed run.**  Until it is cleared, *every* subsequent export
   fails with error -9074 -- including ones for perfectly good files.  One bad input otherwise
   poisons the whole session, which looks exactly like a broken environment.

So a failed export is followed by recovery here: press Escape if macOS Accessibility permission
is available, and force-quit if it is not.  Escape is used rather than clicking a button by
name because the dialog is localised -- the buttons are "Annuleren"/"Herstellen" on a Dutch
install, "Cancel"/"Repair" on an English one -- and matching on either is a bug waiting for the
third language.
"""

from __future__ import annotations

import os
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path

POWERPOINT_APP = Path("/Applications/Microsoft PowerPoint.app")
#: The export script lives in the pptx2svg checkout next to this one.  A checkout elsewhere (a
#: git worktree, say) points at it with ``PPTX2SVG_ORACLE_SCRIPT``.
ORACLE_SCRIPT = Path(
    os.environ.get("PPTX2SVG_ORACLE_SCRIPT")
    or Path(__file__).parents[2] / "pptx2svg" / "tools" / "powerpoint_export_pdf.applescript"
)

#: PowerPoint's sandbox refuses paths outside the user's home with error -9074.
SANDBOX_ROOT = Path(os.path.expanduser("~"))


def available() -> bool:
    return POWERPOINT_APP.exists() and ORACLE_SCRIPT.exists()


@dataclass
class ExportResult:
    ok: bool
    #: "exported", "rejected" (a modal dialog), or "error"
    outcome: str
    detail: str = ""
    recovered: bool = False

    def __bool__(self) -> bool:
        return self.ok


def export_pdf(deck: Path, pdf: Path, *, timeout: int = 90, retries: int = 1) -> ExportResult:
    """Ask PowerPoint to open ``deck`` and export it to ``pdf``.

    A successful result means PowerPoint accepted the file unprompted.  A ``"rejected"``
    outcome means it raised a dialog -- which for this purpose *is* the verdict: the file is
    not one PowerPoint will open cleanly.

    ``retries`` covers the case where the *previous* run left PowerPoint wedged: recovery
    restarts the application, and the first export after a restart can still fail while it
    finishes coming up.  Pass ``retries=0`` when the file is expected to be rejected, so a
    known-bad input is not opened twice.
    """
    if not available():
        raise RuntimeError("PowerPoint or the oracle script is not available")
    for path in (deck, pdf):
        if SANDBOX_ROOT not in path.resolve().parents:
            raise ValueError(f"{path} is outside {SANDBOX_ROOT}; PowerPoint will refuse it")

    result = _attempt(deck, pdf, timeout)
    attempt = 0
    while not result.ok and result.recovered and attempt < retries:
        attempt += 1
        result = _attempt(deck, pdf, timeout)
    return result


def _attempt(deck: Path, pdf: Path, timeout: int) -> ExportResult:
    pdf.unlink(missing_ok=True)
    try:
        completed = subprocess.run(
            ["osascript", str(ORACLE_SCRIPT), str(deck), str(pdf)],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        detail = (completed.stderr or completed.stdout).strip()
        timed_out = False
    except subprocess.TimeoutExpired:
        detail = f"osascript did not return within {timeout}s"
        timed_out = True

    if not timed_out and pdf.exists() and pdf.stat().st_size > 0:
        return ExportResult(True, "exported")

    # A stall, a -1712 AppleEvent timeout or a -9074 all point the same way: PowerPoint is
    # sitting on a dialog and will keep failing until it is cleared.
    blocked = timed_out or "-1712" in detail or "-9074" in detail
    return ExportResult(
        ok=False,
        outcome="rejected" if blocked else "error",
        detail=detail,
        recovered=recover() if blocked else False,
    )


#: Open a deck and save it again as .pptx -- what PowerPoint keeps of a file, it writes back.
#: As in the export script, the paths become files *outside* the ``tell`` block: coerced
#: inside it, ``open`` hangs until the AppleEvent times out.  The presentation is matched by
#: its full path (and, once saved, by the new one), so a deck the user has open is never
#: touched.
_RESAVE_SCRIPT = """
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set inFile to POSIX file inputPath
  set outFile to POSIX file outputPath
  with timeout of 600 seconds
    tell application "Microsoft PowerPoint"
      activate
      open inFile
      set opened to missing value
      repeat with i from 1 to count of presentations
        if (full name of presentation i as text) is inputPath then set opened to presentation i
      end repeat
      if opened is missing value then error "no presentation matched " & inputPath
      if (item 3 of argv) is "potx" then
        save opened in outFile as save as Open XML template
      else
        save opened in outFile as save as Open XML presentation
      end if
      repeat with i from (count of presentations) to 1 by -1
        set candidate to presentation i
        if (full name of candidate as text) is in {inputPath, outputPath} then
          close candidate saving no
        end if
      end repeat
    end tell
  end timeout
end run
"""


def save_as_pptx(deck: Path, out: Path, *, timeout: int = 120, retries: int = 1) -> ExportResult:
    """Have PowerPoint open ``deck`` and save it as ``out`` (.pptx), unprompted."""
    return save_as(deck, out, "pptx", timeout=timeout, retries=retries)


def save_as_potx(deck: Path, out: Path, *, timeout: int = 120, retries: int = 1) -> ExportResult:
    """Have PowerPoint open ``deck`` and save it as a template, ``out`` (.potx)."""
    return save_as(deck, out, "potx", timeout=timeout, retries=retries)


def save_as(deck: Path, out: Path, kind: str, *, timeout: int = 120,
            retries: int = 1) -> ExportResult:
    if not POWERPOINT_APP.exists():
        raise RuntimeError("PowerPoint is not available")
    for path in (deck, out):
        if SANDBOX_ROOT not in path.resolve().parents:
            raise ValueError(f"{path} is outside {SANDBOX_ROOT}; PowerPoint will refuse it")
    result = _resave(deck, out, kind, timeout)
    attempt = 0
    while not result.ok and result.recovered and attempt < retries:
        attempt += 1
        result = _resave(deck, out, kind, timeout)
    return result


def _resave(deck: Path, out: Path, kind: str, timeout: int) -> ExportResult:
    out.unlink(missing_ok=True)
    try:
        completed = subprocess.run(
            ["osascript", "-e", _RESAVE_SCRIPT, str(deck), str(out), kind],
            capture_output=True, text=True, timeout=timeout,
        )
        detail, timed_out = (completed.stderr or completed.stdout).strip(), False
    except subprocess.TimeoutExpired:
        detail, timed_out = f"osascript did not return within {timeout}s", True
    if not timed_out and out.exists() and out.stat().st_size > 0:
        return ExportResult(True, "exported")
    blocked = timed_out or "-1712" in detail or "-9074" in detail
    return ExportResult(False, "rejected" if blocked else "error", detail,
                        recover() if blocked else False)


#: PowerPoint's own new deck: File > New (``make new presentation``), optionally switched to
#: 4:3, with one slide per layout and "S<slide>P<shape>" typed into every placeholder that
#: takes text -- the reference a new deck made by this library is compared with.  Saved by
#: its own path and closed again; a deck the user has open is never touched.
_NEW_DECK_SCRIPT = """
on run argv
  set outputPath to item 1 of argv
  set outFile to POSIX file outputPath
  with timeout of 300 seconds
    tell application "Microsoft PowerPoint"
      set p to make new presentation
      if (item 2 of argv) is "4:3" then
        set slide size of page setup of p to slide size on screen
      end if
      set kinds to {slide layout title slide, slide layout object, slide layout section header}
      set kinds to kinds & {slide layout two objects, slide layout comparison}
      set kinds to kinds & {slide layout title only, slide layout blank}
      set kinds to kinds & {slide layout content with caption, slide layout picture with caption}
      set kinds to kinds & {slide layout vertical text, slide layout vertical title and text}
      repeat with k in kinds
        make new slide at end of p with properties {layout:k}
      end repeat
      repeat with i from 1 to count of slides of p
        set s to slide i of p
        repeat with j from 1 to count of shapes of s
          try
            set content of text range of text frame of shape j of s to "S" & i & "P" & j
          end try
        end repeat
      end repeat
      save p in outFile as save as Open XML presentation
      repeat with i from (count of presentations) to 1 by -1
        if (full name of presentation i as text) is outputPath then close presentation i saving no
      end repeat
    end tell
  end timeout
end run
"""


def new_powerpoint_deck(out: Path, size: str = "16:9", *, timeout: int = 300) -> ExportResult:
    """Have PowerPoint make a new deck, a slide per layout, and save it as ``out``."""
    if not POWERPOINT_APP.exists():
        raise RuntimeError("PowerPoint is not available")
    if SANDBOX_ROOT not in out.resolve().parents:
        raise ValueError(f"{out} is outside {SANDBOX_ROOT}; PowerPoint will refuse it")
    out.unlink(missing_ok=True)
    try:
        completed = subprocess.run(["osascript", "-e", _NEW_DECK_SCRIPT, str(out), size],
                                   capture_output=True, text=True, timeout=timeout)
        detail = (completed.stderr or completed.stdout).strip()
    except subprocess.TimeoutExpired:
        return ExportResult(False, "error", f"osascript did not return within {timeout}s")
    if out.exists() and out.stat().st_size > 0:
        return ExportResult(True, "exported")
    return ExportResult(False, "error", detail)


def recover() -> bool:
    """Clear a stuck PowerPoint.  Returns ``True`` if it looks usable again.

    Note that :func:`responsive` is *not* sufficient evidence on its own: a wedged PowerPoint
    happily answers "count of presentations" with 0 while still refusing every file with
    -9074.  So dismissal is only trusted when Accessibility actually let us press Escape, and
    the fallback is a restart, which is the one thing observed to work every time.
    """
    if dismiss_dialog() and responsive():
        return True

    # Restarting is the fallback when Accessibility is not granted.  It is genuinely worse:
    # PowerPoint reopens the document it was killed over, which raises the dialog again, so
    # this clears the wedge only until the next export.  Granting Accessibility is the real fix.
    force_quit()
    if running():
        return False
    # An `open` sent while PowerPoint is still coming up is refused instantly with -9074 --
    # indistinguishable from the wedge we just cleared -- so wait for it to answer first.
    return relaunch()


def relaunch(timeout: int = 45) -> bool:
    """Start PowerPoint and wait until it answers AppleEvents."""
    subprocess.run(
        ["open", "-g", "-a", str(POWERPOINT_APP)], capture_output=True
    )
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if running() and responsive():
            # Answering is necessary but not sufficient -- give the sandbox a beat to settle
            # before the first file is handed over.
            time.sleep(2)
            return True
        time.sleep(1)
    return False


#: "Cancel" as the repair dialog spells it, across the localisations likely to be met.  The
#: dialog observed on a Dutch install is untitled with exactly two buttons, "Annuleren" and
#: "Herstellen" (Repair); Cancel is the one that gives a verdict rather than silently fixing
#: the file.
CANCEL_LABELS = (
    "Annuleren", "Cancel", "Abbrechen", "Annuler", "Cancelar", "Annulla",
    "Avbryt", "Anuluj", "Отмена", "キャンセル", "取消", "취소",
)

_DISMISS_SCRIPT = """
tell application "System Events"
  tell process "Microsoft PowerPoint"
    repeat with w in windows
      repeat with label in {%s}
        try
          click button label of w
          return "clicked"
        end try
      end repeat
      -- Untitled two-button sheet: the repair dialog in a localisation not listed above.
      try
        if (name of w) is "" and (count of buttons of w) is 2 then
          click button 1 of w
          return "clicked"
        end if
      end try
    end repeat
    return "none"
  end tell
end tell
"""


def dismiss_dialog() -> bool:
    """Click Cancel on PowerPoint's repair dialog via System Events.

    Needs macOS Accessibility permission for whatever runs this; without it System Events
    fails with -1728 and this returns ``False`` so the caller falls back to restarting.

    Pressing Escape does *not* work -- it was tried first and the dialog stayed up.  Only an
    actual button click dismisses it, which is why this goes to the trouble of matching the
    label across localisations instead of sending a keystroke.
    """
    labels = ", ".join('"%s"' % label for label in CANCEL_LABELS)
    try:
        completed = subprocess.run(
            ["osascript", "-e", _DISMISS_SCRIPT % labels],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except subprocess.TimeoutExpired:
        return False
    if completed.returncode != 0 or completed.stdout.strip() != "clicked":
        return False
    time.sleep(1)
    return True


def responsive(timeout: int = 15) -> bool:
    """Is PowerPoint answering AppleEvents?"""
    try:
        completed = subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft PowerPoint" to count of presentations'],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return False
    return completed.returncode == 0


def running() -> bool:
    return subprocess.run(["pgrep", "-x", "Microsoft PowerPoint"], capture_output=True).returncode == 0


def force_quit() -> None:
    """Last resort.

    A graceful ``quit`` is refused while the dialog is up -- PowerPoint answers -128, "cancelled
    by user" -- so there is nothing gentler that works.  Safe in this context because the oracle
    never leaves a presentation open: any deck it opened, it also closed.
    """
    subprocess.run(["pkill", "-x", "Microsoft PowerPoint"], capture_output=True)
    for _ in range(10):
        if not running():
            return
        time.sleep(0.5)


def pdf_page_count(pdf: Path) -> int:
    """Pages in a PDF, by counting page objects -- enough for the PDFs PowerPoint writes."""
    import re

    return len(re.findall(rb"/Type\s*/Page(?![a-zA-Z])", pdf.read_bytes()))


def pdf_pages(pdf: Path, dpi: int = 72):
    """The PDF's pages as PIL images, or ``None`` without pypdfium2 and Pillow installed."""
    try:
        import pypdfium2
    except ImportError:
        return None
    document = pypdfium2.PdfDocument(str(pdf))
    try:
        return [page.render(scale=dpi / 72).to_pil().convert("RGB") for page in document]
    finally:
        document.close()


def pdf_texts(pdf: Path) -> list[str] | None:
    """Each page's text, line breaks removed, or ``None`` without pypdfium2 installed."""
    try:
        import pypdfium2
    except ImportError:
        return None
    document = pypdfium2.PdfDocument(str(pdf))
    try:
        texts = []
        for page in document:
            text = page.get_textpage().get_text_range()
            texts.append(text.replace("\r", "").replace("\n", ""))
        return texts
    finally:
        document.close()


def cleanup(*paths: Path) -> None:
    """Remove decks, PDFs and the ``~$`` lock files PowerPoint leaves behind on a bad open."""
    for path in paths:
        path.unlink(missing_ok=True)
        (path.parent / f"~${path.name}").unlink(missing_ok=True)
