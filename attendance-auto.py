#!/usr/bin/env python3
"""
attendance_sync.py - Google Form attendance responses -> roster grid in Google Sheets.

SET_UP Explanation
Write Testcode... NPM run? Docker image? How to maintain the dependencies?

LOCAL EXCEL MODE (no Google account or API needed)
 1. pip install openpyxl python-dateutil
 2. python attendance-auto.py --xlsx attendance.xlsx
      The first run creates the workbook with two tabs - fill them in:
        Attendance : Last Name | First Name
        Codes      : Date | Code, one row per class day MM/DD/YYYY (e.g. 11/3/2026)
 3. After class (preferrably every week): Attendance Google Form > Responses > "Download responses (.csv)"
 4. Move the files in the same folder (including the python code)
        python attendance_sync.py --xlsx attendance.xlsx --responses responses.csv
      Close the workbook in Excel first. A backup (attendance.xlsx.bak) is saved each run.
      Each export contains every response so far, so just re-export and re-run.

GOOGLE SHEETS MODE - one-time setup
 1. Google Cloud Console: create a project, enable the "Google Sheets API" and
    "Google Drive API", create a Service Account, and download its JSON key as
    service_account.json (next to this script).
 2. Open the Google Sheet your Form is linked to (Form > Responses > "Link to Sheets")
    and Share it with the service account's email address (Editor).
 3. In that same spreadsheet add these tabs:
      Attendance : row 1 = Last Name | First Name | Aliases
                   then one student per row from row 2 down (copy from Blackboard).
                   Date columns are added to the right automatically.
                   Aliases is optional: nicknames / alternate spellings, separated by ";"
                   e.g.  Mike Johnson; Mikey Johnson
      Codes      : row 1 = Date | Code
                   one row per class day, e.g.  11/3/2026 | MAPLE42
      (a "Review" tab is created automatically)
 4. pip install "gspread>=6" python-dateutil
 5. Paste the spreadsheet ID (the long string in the sheet's URL) into SPREADSHEET_ID.

USAGE
    python attendance_sync.py                                   # Google mode: update the sheet
    python attendance_sync.py --responses responses.csv         # Google mode, using an exported CSV
    python attendance_sync.py --xlsx attendance.xlsx --responses responses.csv   # local Excel mode
    add --dry-run to any of these to print what would happen and write nothing

Safe to re-run any time: it recomputes every class day from the responses.
 - A ✅ is never downgraded to ❌ (so you can fix things by hand).
 - Anything else you type in a date cell (e.g. "E" for excused) is left alone.
 - Today's column only gets ✅ marks; ❌ are filled in from the next day on.
"""
import argparse
import csv
import os
import re
import shutil
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from zoneinfo import ZoneInfo

from dateutil import parser as dateparser

warnings.filterwarnings("ignore")

# ----------------------------------------------------------------- CONFIG
SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE"
RESPONSES_TAB = "Form Responses 1"
ATTENDANCE_TAB = "Attendance"
CODES_TAB = "Codes"
REVIEW_TAB = "Review"
TIMEZONE = "America/New_York"

PRESENT, ABSENT = "✅", "❌"
FIRST_DATE_COL = 3          # 0-based: A=Last Name, B=First Name, C=Aliases, dates start at D
AUTO_FUZZY = 0.92           # similarity needed to auto-accept a misspelled name
SUGGEST_FUZZY = 0.75        # similarity needed to show a suggestion in the Review tab


# ----------------------------------------------------------- name helpers
def norm(text):
    """lowercase, strip accents/punctuation: "Seán O'Brien" -> "sean o brien"."""
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def name_key(tokens):
    """Order-insensitive key, so 'Smith John' == 'John Smith'."""
    return " ".join(sorted(tokens))


def norm_code(text):
    """"""
    return re.sub(r"\s+", "", text or "").casefold()


def to_date(text):
    """"""
    text = (text or "").strip()
    if not text:
        return None
    try:
        return dateparser.parse(text, dayfirst=False).date()
    except (ValueError, OverflowError):
        return None


@dataclass
class Student:
    row: int            # index into the Attendance values (sheet row = row + 1)
    last: str
    first: str
    aliases: str = ""

    def __post_init__(self):
        tokens = norm(f"{self.first} {self.last}").split()
        self.tokens = set(tokens)
        self.key = name_key(tokens)
        self.alias_keys = {
            name_key(norm(a).split()) for a in re.split(r"[;\n]", self.aliases) if norm(a)
        }

    @property
    def display(self):
        return f"{self.last}, {self.first} (row {self.row + 1})"


class Matcher:
    """Tries progressively looser matches; anything uncertain goes to the Review tab."""

    def __init__(self, students):
        self.students = students

    @staticmethod
    def _pick(hits, kind):
        if len(hits) == 1:
            return hits[0], kind, ""
        return None, "ambiguous", "; ".join(h.display for h in hits)

    def match(self, first, last):
        tokens = norm(f"{first} {last}").split()
        if not tokens:
            return None, "empty", ""
        key, tset = name_key(tokens), set(tokens)

        # 1. alias you added (nicknames etc.)
        hits = [s for s in self.students if key in s.alias_keys]
        if hits:
            return self._pick(hits, "alias")
        # 2. exact name (ignores case, accents, punctuation, first/last swap)
        hits = [s for s in self.students if s.key == key]
        if hits:
            return self._pick(hits, "exact")
        # 3. middle name added/omitted: one name's words contain the other's (2+ words shared)
        hits = [
            s for s in self.students
            if len(tset & s.tokens) >= 2 and (tset <= s.tokens or s.tokens <= tset)
        ]
        if hits:
            return self._pick(hits, "middle-name")
        # 4. typo-level similarity
        scored = sorted(((SequenceMatcher(None, key, s.key).ratio(), s) for s in self.students),
                        key=lambda x: x[0], reverse=True)
        if scored:
            best, student = scored[0]
            second = scored[1][0] if len(scored) > 1 else 0
            if best >= AUTO_FUZZY and best - second >= 0.05:
                return student, "fuzzy", f"{best:.0%}"
            if best >= SUGGEST_FUZZY:
                return None, "suggest", f"{student.display}, {best:.0%} similar"
        # 5. weaker hint for the Review tab only: shares 2+ name parts with exactly one student
        #    (e.g. an extra surname the roster doesn't have)
        overlap = [s for s in self.students if len(tset & s.tokens) >= 2]
        if len(overlap) == 1:
            return None, "suggest", f"{overlap[0].display}, shares {len(tset & overlap[0].tokens)} name parts"
        return None, "none", ""


# ------------------------------------------------------------ core logic
def find_col(headers, keyword):
    hits = [i for i, h in enumerate(headers) if keyword in h.lower()]
    if len(hits) != 1:
        raise SystemExit(f"Expected exactly one form column containing '{keyword}', "
                         f"found {len(hits)}. Headers are: {headers}")
    return hits[0]


def load_codes(values):
    codes = {}
    for row in values[1:]:
        if len(row) >= 2:
            d, code = to_date(row[0]), norm_code(row[1])
            if d and code:
                codes[d] = code
    return codes


def load_students(att_values):
    students = []
    for i, row in enumerate(att_values):
        if i == 0:
            continue
        row = row + [""] * (FIRST_DATE_COL - len(row))
        last, first = row[0].strip(), row[1].strip()
        if last or first:
            students.append(Student(i, last, first, row[2]))
    return students


def process(att_values, codes, resp_values, today):
    """Pure logic (no network). Returns dict with the new date block and review rows."""
    if not att_values:
        raise SystemExit("The Attendance tab is empty (need: Last Name | First Name | Aliases).")
    students = load_students(att_values)
    if not students:
        raise SystemExit("No students found on the Attendance tab (row 2 down).")
    matcher = Matcher(students)

    headers, rows = resp_values[0], resp_values[1:]
    c_ts, c_code, c_first, c_last = (find_col(headers, k) for k in ("timestamp", "code", "first", "last"))

    present = defaultdict(set)      # class date -> student rows marked present
    flags = {}                      # dedup key -> (date, student_row_or_None, review row)
    stats = defaultdict(int)

    for r in rows:
        r = r + [""] * (len(headers) - len(r))
        day = to_date(r[c_ts])
        if day is None:
            stats["unreadable timestamp"] += 1
            continue
        if day not in codes:
            stats["submitted on a day with no code"] += 1
            continue
        first, last, typed = r[c_first].strip(), r[c_last].strip(), r[c_code].strip()
        student, kind, detail = matcher.match(first, last)
        base = [day.isoformat(), r[c_ts], first, last]
        fkey = (day, norm(first + " " + last))

        if norm_code(typed) != codes[day]:
            flags[fkey + ("code",)] = (day, student.row if student else None,
                                       base + [f"Wrong code (typed: {typed})", student.display if student else ""])
        elif student is not None:
            present[day].add(student.row)
            if kind == "fuzzy":
                flags[fkey + ("fuzzy",)] = (day, None, base + ["Auto-matched by similarity - verify",
                                                               f"{student.display}, {detail}"])
        else:
            issue = {"ambiguous": "Ambiguous - several roster matches",
                     "suggest": "No exact match - possible match",
                     "none": "No roster match",
                     "empty": "Blank name"}[kind]
            flags[fkey + ("name",)] = (day, None, base + [issue, detail])

    # A wrong-code entry is noise if that student also submitted the right code that day.
    review = [v for (d, srow, v) in flags.values() if srow is None or srow not in present[d]]
    review.sort(key=lambda v: (v[0], v[2], v[3]))

    # ---- build the new date block (columns D onward)
    width = max(FIRST_DATE_COL, max(len(r) for r in att_values))
    last_row = max(s.row for s in students)
    grid = [list(att_values[i]) + [""] * (width - len(att_values[i])) for i in range(last_row + 1)]

    col_of = {}
    for j in range(FIRST_DATE_COL, width):
        d = to_date(grid[0][j])
        if d:
            col_of[d] = j
    class_dates = sorted(d for d in codes if d <= today)
    new_dates = [d for d in class_dates if d not in col_of]
    for d in new_dates:
        col_of[d] = width
        width += 1
        for row in grid:
            row.append("")
        grid[0][col_of[d]] = d.isoformat()

    for s in students:
        row = grid[s.row]
        for d in class_dates:
            j = col_of[d]
            existing = row[j].strip()
            if existing == PRESENT or s.row in present[d]:
                row[j] = PRESENT
            elif existing in ("", ABSENT):
                row[j] = ABSENT if d < today else row[j]
            # anything else is a manual note ("E", "late", ...): leave it

    summary = {d: (len(present[d]), len(students)) for d in class_dates}
    return {"block": [row[FIRST_DATE_COL:] for row in grid], "review": review,
            "summary": summary, "new_dates": new_dates, "stats": dict(stats)}


# -------------------------------------------------------------------- I/O
def write_review(sh, rows):
    import gspread
    header = REVIEW_HEADER
    values = [header] + rows
    try:
        ws = sh.worksheet(REVIEW_TAB)
        ws.clear()
        ws.resize(rows=max(ws.row_count, len(values) + 20), cols=max(ws.col_count, len(header)))
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=REVIEW_TAB, rows=len(values) + 20, cols=len(header))
    ws.update(values=values, range_name="A1", value_input_option="RAW")


REVIEW_HEADER = ["Class date", "Submitted at", "First name", "Last name", "Issue", "Details / suggestion"]


def cell_str(v):
    """openpyxl cell value -> the plain string Google Sheets would hand us."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if v.time() == datetime.min.time() else v.isoformat(sep=" ")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def sheet_values(ws):
    """Worksheet -> list of rows of strings, trailing blank cells/rows trimmed."""
    rows = [[cell_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
    for r in rows:
        while r and not r[-1].strip():
            r.pop()
    while rows and not rows[-1]:
        rows.pop()
    return rows


def read_table(path):
    """Exported responses: .csv, or .xlsx (first sheet)."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        return sheet_values(openpyxl.load_workbook(path, data_only=True).worksheets[0])
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def report(result, dry_run):
    for d, (n, total) in result["summary"].items():
        print(f"{d}: {n}/{total} present")
    for k, v in result["stats"].items():
        print(f"skipped {v} submission(s): {k}")
    print(f"{len(result['review'])} item(s) for the Review tab")
    if dry_run:
        for row in result["review"]:
            print("  ", row)
    return dry_run


def run_local(args, today):
    import openpyxl
    from openpyxl.utils import get_column_letter

    path = args.xlsx
    if not os.path.exists(path):
        wb = openpyxl.Workbook()
        wb.active.title = ATTENDANCE_TAB
        wb.active.append(["Last Name", "First Name", "Aliases"])
        wb.create_sheet(CODES_TAB).append(["Date", "Code"])
        wb.save(path)
        raise SystemExit(f"Created {path} with '{ATTENDANCE_TAB}' and '{CODES_TAB}' tabs. "
                         "Fill in the roster and the class-day codes, then run again with --responses.")
    if not args.responses:
        raise SystemExit("Local mode needs --responses <exported CSV or XLSX of the Form responses>.")

    wb = openpyxl.load_workbook(path)
    for tab in (ATTENDANCE_TAB, CODES_TAB):
        if tab not in wb.sheetnames:
            raise SystemExit(f"{path} has no '{tab}' tab (found: {wb.sheetnames}).")
    ws = wb[ATTENDANCE_TAB]

    result = process(sheet_values(ws), load_codes(sheet_values(wb[CODES_TAB])),
                     read_table(args.responses), today)
    if report(result, args.dry_run):
        return

    # new class-day headers become real Excel dates; everything else in the header row is untouched
    for i, row in enumerate(result["block"]):
        for j, v in enumerate(row):
            col = FIRST_DATE_COL + 1 + j
            if i == 0:
                d = to_date(v) if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v or "") else None
                if d in result["new_dates"]:
                    cell = ws.cell(row=1, column=col, value=d)
                    cell.number_format = "mmm d"
                    ws.column_dimensions[get_column_letter(col)].width = 9
            else:
                ws.cell(row=i + 1, column=col, value=v or None)

    if REVIEW_TAB in wb.sheetnames:
        del wb[REVIEW_TAB]
    rv = wb.create_sheet(REVIEW_TAB)
    rv.append(REVIEW_HEADER)
    for row in result["review"]:
        rv.append(row)
    for letter, w in zip("ABCDEF", (12, 22, 14, 16, 36, 52)):
        rv.column_dimensions[letter].width = w

    shutil.copy2(path, path + ".bak")
    try:
        wb.save(path)
    except PermissionError:
        raise SystemExit(f"Could not write {path}. Close it in Excel and run again.")
    print(f"Saved {path} (previous version kept as {path}.bak).")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", help="local Excel workbook to use instead of Google Sheets")
    ap.add_argument("--responses", "--csv", dest="responses",
                    help="exported Form responses (.csv or .xlsx); Google mode reads the linked tab if omitted")
    ap.add_argument("--dry-run", action="store_true", help="show results, write nothing")
    args = ap.parse_args()

    today = datetime.now(ZoneInfo(TIMEZONE)).date()
    if args.xlsx:
        return run_local(args, today)

    import gspread
    from gspread.utils import rowcol_to_a1

    sh = gspread.service_account(filename=SERVICE_ACCOUNT_FILE).open_by_key(SPREADSHEET_ID)
    att_ws = sh.worksheet(ATTENDANCE_TAB)

    if args.responses:
        resp_values = read_table(args.responses)
    else:
        resp_values = sh.worksheet(RESPONSES_TAB).get_all_values()

    result = process(att_ws.get_all_values(), load_codes(sh.worksheet(CODES_TAB).get_all_values()),
                     resp_values, today)

    if report(result, args.dry_run):
        return

    block = result["block"]
    if block and block[0]:
        needed = FIRST_DATE_COL + len(block[0])
        if att_ws.col_count < needed:
            att_ws.add_cols(needed - att_ws.col_count)
        att_ws.update(values=block, range_name=rowcol_to_a1(1, FIRST_DATE_COL + 1),
                      value_input_option="RAW")
    write_review(sh, result["review"])
    print("Done.")


if __name__ == "__main__":
    main()