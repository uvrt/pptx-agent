"""What PowerPoint writes for a table's row heights and its frame's height (round 54, R114).

A row's stored height (``a:tr@h``) is a minimum: PowerPoint draws a row whose text needs more
taller (``tools/table_rows_probe.py``).  This probe asks what PowerPoint *writes*: a made-up
summary table -- two columns, a number and a text, "No Style, No Grid", 18 pt -- four times
on a slide each, rows stored 62 pt (as rows copied from a 62 pt sample row) or 0, saved by
PowerPoint after

- opening and showing each slide (what it writes for ``a:tr@h`` and the frame's ``cy``);
- replacing cell text through its object model (a short text where a long one was, and the
  other way round) -- the screen is often locked when the oracle runs, so typing by
  keystrokes is not measured;
- asking for row heights below the text's (0 pt, and 62 pt / 10 pt on a table of 0 pt rows),
  as dragging a row to its smallest does.

::

    python tools/table_fit_probe.py build ~/table-fit-probe.pptx
    python tools/table_fit_probe.py drive ~/table-fit-probe.pptx ~/table-fit-out.pptx
    python tools/table_fit_probe.py read ~/table-fit-out.pptx

``drive`` prints the row heights PowerPoint reports for each slide as it goes; ``read`` the
heights and frame height the saved deck stores.  Inputs and outputs go in the home folder's
root (PowerPoint's sandbox).  The measurements are in ``tests/fixtures/table-fit-probe.json``;
``tests/test_fit_rows.py`` holds ``Table.fit_rows``, ``row_text_heights`` and the frame's
height to them.
"""

from __future__ import annotations

import re
import subprocess
import sys
import time
import zipfile

PT = 12700
TEXTS = ["Shelf count closed",
         "Move the spare parts to the north shelf, label every bin with its new number, and "
         "check the counts against last month's sheet before Friday",
         "Order two more ladders for the back room and return the broken one",
         "Done",
         "Shelf count closed"]

_DRIVE = """
on rowsOf(sl)
  tell application "Microsoft PowerPoint"
    set sh to shape 1 of sl
    set t to table object of sh
    set acc to "frame " & (height of sh) & " rows"
    repeat with i from 1 to count of rows of t
      set acc to acc & " " & (height of row i of t)
    end repeat
    return acc
  end tell
end rowsOf

on setText(sl, r, c, s)
  tell application "Microsoft PowerPoint"
    set theCell to get cell from (table object of shape 1 of sl) row r column c
    set content of text range of text frame of (shape of theCell) to s
  end tell
end setText

on showSlide(n)
  tell application "Microsoft PowerPoint"
    go to slide (view of document window 1) number n
  end tell
  delay 1
end showSlide

on run argv
  set log_ to ""
  -- The file references are made outside PowerPoint's tell block, as the PDF export script
  -- makes them: made inside it, the open timed out.
  set inFile to POSIX file (item 1 of argv)
  set outFile to POSIX file (item 2 of argv)
  with timeout of 120 seconds
    tell application "Microsoft PowerPoint"
      open inFile
      set p to active presentation
      repeat with i from 1 to count of slides of p
        my showSlide(i)
        set log_ to log_ & "opened s" & i & ": " & my rowsOf(slide i of p) & linefeed
      end repeat
      my setText(slide 3 of p, 1, 2, item 3 of argv)
      my setText(slide 3 of p, 2, 2, "Short now")
      my setText(slide 4 of p, 1, 2, item 3 of argv)
      my setText(slide 4 of p, 2, 2, "Short now")
      set t to table object of shape 1 of slide 1 of p
      repeat with i from 1 to count of rows of t
        set height of row i of t to 0
      end repeat
      set t to table object of shape 1 of slide 2 of p
      set height of row 1 of t to 62
      set height of row 2 of t to 10
      repeat with i from 1 to count of slides of p
        my showSlide(i)
        set log_ to log_ & "edited s" & i & ": " & my rowsOf(slide i of p) & linefeed
      end repeat
      save p in outFile as save as Open XML presentation
      close p saving no
    end tell
  end timeout
  return log_
end run
"""


def build(deck_path: str) -> None:
    from pptx_agent import Document

    deck = Document.new(size="16:9")
    for heights in ([62 * PT] * 5, [0] * 5, [62 * PT] * 5, [0] * 5):
        slide = deck.add_slide("Blank")
        table = slide.add_table(5, 2, 60 * PT, 60 * PT, 470 * PT, sum(heights),
                                style="No Style, No Grid").table
        table.set_column_width(0, 50 * PT)
        table.set_column_width(1, 420 * PT)
        table.set_row_heights(heights)
        for row, text in enumerate(TEXTS):
            table.cell(row, 0).text = str(row + 1)
            table.cell(row, 1).text = text
        print(f"s{len(deck.slides)}: stored {[h / PT for h in table.row_heights]}, drawn "
              f"{[round(h / PT, 1) for h in table.drawn_row_heights]}")
    deck.save(deck_path)


def drive(deck_path: str, out_path: str) -> None:
    # A deck sent while PowerPoint is still starting times out: start it and wait until it
    # answers first (as tests/oracle.py's relaunch does).
    subprocess.run(["open", "-g", "-a", "Microsoft PowerPoint"], capture_output=True)
    for _ in range(45):
        asked = subprocess.run(["osascript", "-e", 'tell application "Microsoft PowerPoint" '
                                "to count of presentations"], capture_output=True, timeout=30)
        if asked.returncode == 0:
            break
        time.sleep(1)
    time.sleep(2)
    done = subprocess.run(["osascript", "-e", _DRIVE, deck_path, out_path, TEXTS[1]],
                          capture_output=True, text=True, timeout=300)
    print(done.stdout or done.stderr)


def read(deck_path: str) -> None:
    with zipfile.ZipFile(deck_path) as package:
        names = sorted((n for n in package.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                       key=lambda n: int(re.findall(r"\d+", n)[0]))
        for name in names:
            xml = package.read(name).decode()
            cy = re.search(r'<p:graphicFrame>.*?<a:ext cx="\d+" cy="(\d+)"', xml, re.S)
            rows = [int(h) / PT for h in re.findall(r'<a:tr h="(\d+)"', xml)]
            print(f"{name}: frame {int(cy.group(1)) / PT} rows {rows}")


if __name__ == "__main__":
    {"build": build, "drive": drive, "read": read}[sys.argv[1]](*sys.argv[2:])
