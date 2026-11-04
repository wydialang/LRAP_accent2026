"""Build an updated speaker spreadsheet that records which speakers have a TextGrid.

Reads the Speech Accent Archive speaker spreadsheet (by default the November 2023
one), scans the textgrids/ folder for TextGrid files and their paired .wav files,
and writes a new workbook.

The new workbook's first sheet is the original sheet with four columns added on
the right:

    has_textgrid       yes / no
    textgrid_path      path relative to the archive folder ("; "-separated if several)
    in_original_sheet  yes for original rows, no for rows this script added
    review_flag        anything about the row worth a human look (blank if nothing)

Every TextGrid that matches no spreadsheet row gets an added row at the bottom. Its
speaker metadata is unknown, so those cells are filled with MISSING and the row is
highlighted. review_flag suggests existing rows it may belong to.

A second sheet, textgrid_log, lists every issue found: unmatched TextGrids,
TextGrids whose name differs from speech_sample, duplicates, and unpaired files.

TextGrids are matched to rows by speech_sample (minus ".mp3"), in this order:
    1. TEXTGRID_ALIASES below, for files known to be named differently
    2. exact name match
    3. name match ignoring case and punctuation (sa_a1 -> sa'a1)

Usage:
    python update_speaker_spreadsheet.py [--archive-dir DIR] [--spreadsheet FILE] [--output FILE]

Requires openpyxl.
"""

import argparse
import difflib
import re
import sys
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Script lives in LRAP_accent2026/code/speechaccentarchive/; the archive is LRAP_accent2026/Speech accent archive.
ARCHIVE_DIR = Path(__file__).resolve().parents[2] / "Speech accent archive"
DEFAULT_SPREADSHEET_NAME = "speakers-1 november 2023.xlsx"

# TextGrid name -> the speech_sample name it belongs to (both without extension),
# for TextGrids named differently from the spreadsheet. Add entries here once a
# suggested match in the review_flag column is confirmed, then re-run.
# sinhalese1-4: each paired .wav is byte-identical to processed_wav_files/sinhalaN.wav.
TEXTGRID_ALIASES = {
    "sinhalese1": "sinhala1",
    "sinhalese2": "sinhala2",
    "sinhalese3": "sinhala3",
    "sinhalese4": "sinhala4",
}

# Written into the unknown metadata cells of added rows. None leaves them blank,
# which pandas reads as NaN; set e.g. "UNKNOWN" to show visible text instead.
MISSING = None

NEW_COLUMNS = {  # header -> column width
    "has_textgrid": 13,
    "textgrid_path": 55,
    "in_original_sheet": 18,
    "review_flag": 100,
}
REQUIRED_COLUMNS = ["speakerid", "speaker", "alternative_native_language", "speech_sample"]
ADDED_ROW_FILL = PatternFill("solid", fgColor="FFE699")
LOG_COLUMNS = {"issue": 34, "textgrid": 18, "spreadsheet_row": 30, "details": 80, "files": 70}


@dataclass(eq=False)
class SpeakerRow:
    excel_row: int
    sample: str | None  # speech_sample as written, e.g. "afrikaans1.mp3"
    stem: str | None  # speech_sample without extension, e.g. "afrikaans1"
    speakerid: object
    speaker: object
    alt_language: str | None
    textgrids: list[str] = field(default_factory=list)
    textgrid_names: set[str] = field(default_factory=set)
    flags: list[str] = field(default_factory=list)

    def label(self):
        return f"{self.sample} (row {self.excel_row})"


def norm(name):
    """Lowercase and keep only letters and digits, so "sa'a1" and "sa_a1" compare equal."""
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def read_speakers(ws, col):
    rows = []
    for r in range(2, ws.max_row + 1):
        if all(c.value is None or str(c.value).strip() == "" for c in ws[r]):
            continue
        sample = ws.cell(r, col["speech_sample"]).value
        sample = str(sample).strip() if sample is not None and str(sample).strip() else None
        rows.append(
            SpeakerRow(
                excel_row=r,
                sample=sample,
                stem=Path(sample).stem if sample else None,
                speakerid=ws.cell(r, col["speakerid"]).value,
                speaker=ws.cell(r, col["speaker"]).value,
                alt_language=ws.cell(r, col["alternative_native_language"]).value,
            )
        )
    return rows


def scan_textgrids(textgrid_dir):
    """Find every TextGrid under textgrid_dir.

    Returns ({name: [TextGrid paths]}, {(folder, name) that have a .wav}).
    """
    files = sorted(p for p in textgrid_dir.rglob("*") if p.is_file())
    textgrids = defaultdict(list)
    for p in files:
        if p.suffix.lower() == ".textgrid":
            textgrids[p.stem].append(p)
    wavs = {(p.parent, p.stem) for p in files if p.suffix.lower() == ".wav"}
    return textgrids, wavs


def find_row(name, by_stem, by_norm):
    """Return (row, note) for the spreadsheet row the TextGrid `name` belongs to.

    note explains a match that isn't an exact name match. row is None if nothing matches.
    """
    if name in TEXTGRID_ALIASES:
        target = TEXTGRID_ALIASES[name]
        if target in by_stem:
            return by_stem[target], f"TextGrid is named {name} (matched via TEXTGRID_ALIASES)"
        return None, f"TEXTGRID_ALIASES maps {name} to {target}, which is not in the spreadsheet"
    if name in by_stem:
        return by_stem[name], None
    if norm(name) in by_norm:
        return by_norm[norm(name)], f"TextGrid is named {name} (matched ignoring case/punctuation)"
    return None, None


def suggest_matches(name, rows, n=3):
    """Spreadsheet rows whose name resembles the TextGrid `name`, best first.

    Each row is compared under its speech_sample name and under its alternative
    language name + speaker number, which is how e.g. sinhalese1 relates to sinhala1.
    """
    keys = {}
    for row in rows:
        if row.stem:
            keys.setdefault(norm(row.stem), row)
        if row.alt_language and row.speaker is not None:
            for alt in str(row.alt_language).split(","):
                if norm(alt):
                    keys.setdefault(norm(alt) + str(row.speaker).removesuffix(".0"), row)
    found = []
    for key in difflib.get_close_matches(norm(name), keys, n=n * 3, cutoff=0.8):
        if keys[key] not in found:
            found.append(keys[key])
    return found[:n]


def build(spreadsheet, archive_dir, output):
    wb = load_workbook(spreadsheet)
    ws = wb.worksheets[0]
    header = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    if missing := [c for c in REQUIRED_COLUMNS if c not in header]:
        sys.exit(f"{spreadsheet.name} has no column(s) named {missing}")
    if already := [c for c in NEW_COLUMNS if c in header]:
        sys.exit(f"{spreadsheet.name} already has {already}; run this on the original spreadsheet")

    rows = read_speakers(ws, header)
    log = []  # (issue, textgrid, spreadsheet_row, details, files)

    by_stem = {}
    for row in rows:
        if row.stem is None:
            continue
        if row.stem in by_stem:
            log.append(("speech_sample listed twice", "", row.label(),
                        f"also {by_stem[row.stem].label()}; TextGrids go to the first", ""))
        by_stem.setdefault(row.stem, row)
    norm_counts = Counter(norm(s) for s in by_stem)
    by_norm = {norm(s): r for s, r in by_stem.items() if norm_counts[norm(s)] == 1}

    textgrid_dir = archive_dir / "textgrids"
    textgrids, wavs = scan_textgrids(textgrid_dir)
    rel = lambda p: p.relative_to(archive_dir).as_posix()

    added = []  # (TextGrid name, paths, flags) for TextGrids matching no row
    for name, paths in sorted(textgrids.items()):
        files = "; ".join(rel(p) for p in paths)
        row, match_note = find_row(name, by_stem, by_norm)
        target = row.label() if row else ""
        flags = []

        if match_note:
            flags.append(match_note)
            if row:
                log.append(("name differs from speech_sample", name, target, match_note, files))
        if len(paths) > 1:
            same = len({p.read_bytes() for p in paths}) == 1
            note = f"{name}.TextGrid is in {len(paths)} folders ({'identical copies' if same else 'COPIES DIFFER'})"
            flags.append(note)
            log.append(("TextGrid in several folders", name, target, note, files))
        for p in paths:
            if (p.parent, p.stem) not in wavs:
                flags.append(f"no paired .wav for {rel(p)}")
                log.append(("TextGrid without .wav", name, target, "", rel(p)))

        if row:
            row.textgrids += [rel(p) for p in paths]
            row.textgrid_names.add(name)
            row.flags += flags
            continue

        suggestions = suggest_matches(name, rows)
        if suggestions:
            hint = "possible match: " + "; ".join(f"{r.sample} (speakerid {r.speakerid})" for r in suggestions)
        else:
            hint = "no similar speech_sample in the spreadsheet"
        flags.insert(0, f"NOT IN ORIGINAL SPREADSHEET, speaker metadata unknown; {hint}")
        added.append((name, [rel(p) for p in paths], flags))
        log.append(("not in spreadsheet (row added)", name, "", hint, files))

    for row in rows:
        if len(row.textgrid_names) > 1:
            note = "several TextGrids match this row: " + ", ".join(sorted(row.textgrid_names))
            row.flags.append(note)
            log.append(("several TextGrids for one row", "", row.label(), note, "; ".join(row.textgrids)))

    textgrid_places = {(p.parent, p.stem) for ps in textgrids.values() for p in ps}
    for folder, name in sorted(wavs - textgrid_places):
        log.append((".wav without TextGrid", name, "", "", rel(folder / f"{name}.wav")))

    write_workbook(wb, ws, header, rows, added, log, output)

    n_tg_files = sum(len(p) for p in textgrids.values())
    print(f"Read {len(rows)} rows from {spreadsheet.name}")
    print(f"Found {len(textgrids)} distinct TextGrids ({n_tg_files} files) in {textgrid_dir}")
    print(f"  {sum(1 for r in rows if r.textgrids)} spreadsheet rows have a TextGrid")
    print(f"  {len(added)} TextGrids match no row and were added as new rows")
    print(f"  {len(log)} issues logged:")
    for issue, name, target, details, _ in log:
        print(f"    - {issue}: {name or target}" + (f" -> {target}" if name and target else "")
              + (f" | {details}" if details else ""))
    print(f"Wrote {output}")


def write_workbook(wb, ws, header, rows, added, log, output):
    first_new = max(header.values()) + 1
    header_font = ws.cell(1, 1).font
    for i, (name, width) in enumerate(NEW_COLUMNS.items()):
        ws.cell(1, first_new + i, name).font = copy(header_font)
        ws.column_dimensions[get_column_letter(first_new + i)].width = width

    def write_extra(r, paths, in_original, flags, font):
        values = ["yes" if paths else "no", "; ".join(paths) or None,
                  "yes" if in_original else "no", "; ".join(flags) or None]
        for i, value in enumerate(values):
            ws.cell(r, first_new + i, value).font = copy(font)

    for row in rows:
        write_extra(row.excel_row, row.textgrids, True, row.flags, ws.cell(row.excel_row, 1).font)

    body_font = ws.cell(2, 1).font
    for name, paths, flags in added:
        r = ws.max_row + 1
        for c in range(1, first_new):
            cell = ws.cell(r, c, MISSING)
            cell.font = copy(body_font)
        write_extra(r, paths, False, flags, body_font)
        for c in range(1, first_new + len(NEW_COLUMNS)):
            ws.cell(r, c).fill = ADDED_ROW_FILL

    log_ws = wb.create_sheet("textgrid_log")
    log_ws.append(list(LOG_COLUMNS))
    for i, width in enumerate(LOG_COLUMNS.values(), start=1):
        log_ws.cell(1, i).font = Font(bold=True)
        log_ws.column_dimensions[get_column_letter(i)].width = width
    for entry in log:
        log_ws.append(list(entry))
    log_ws.freeze_panes = "A2"

    wb.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--archive-dir", type=Path, default=ARCHIVE_DIR,
                        help="Speech accent archive folder containing textgrids/ and excel_spreadsheets/")
    parser.add_argument("--spreadsheet", type=Path,
                        help=f"input speaker spreadsheet (default: excel_spreadsheets/{DEFAULT_SPREADSHEET_NAME})")
    parser.add_argument("--output", type=Path,
                        help="output .xlsx (default: excel_spreadsheets/speakers_textgrids_<today>.xlsx)")
    args = parser.parse_args()

    archive_dir = args.archive_dir.resolve()
    spreadsheets = archive_dir / "excel_spreadsheets"
    spreadsheet = args.spreadsheet or spreadsheets / DEFAULT_SPREADSHEET_NAME
    output = args.output or spreadsheets / f"speakers_textgrids_{date.today():%Y-%m-%d}.xlsx"
    if output.resolve() == spreadsheet.resolve():
        sys.exit("--output must differ from --spreadsheet; the original is never overwritten")
    build(spreadsheet, archive_dir, output)


if __name__ == "__main__":
    main()
