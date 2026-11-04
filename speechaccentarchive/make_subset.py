"""Pick a small speaker subset for pilot runs and write the file list rclone needs.

Picks N native (L1) English speakers and N non-native (L2) speakers, using only
speakers that have a TextGrid. L1 speakers are split evenly by gender. Each L2
speaker comes from a different native language, taken from the most common
languages in the archive, so the subset can be grown later without changing who
is in it. The same --seed always gives the same speakers.

Writes two files to --out-dir:

    <name>_speakers.csv   one row per chosen speaker (group, speaker, metadata, paths)
    <name>_files.txt      TextGrid + wav paths relative to the archive folder,
                          for `rclone copy --files-from`

Usage:
    python make_subset.py [--archive-dir DIR] [--n 10] [--seed 0] [--name pilot_20]
                          [--spreadsheet FILE] [--out-dir DIR]

--archive-dir is the Speech Accent Archive folder (default: the one in this project).
The spreadsheet is read from its excel_spreadsheets/ folder unless --spreadsheet is
given. On Savio, point it at the scratch copy:

    python make_subset.py --archive-dir /global/scratch/users/$USER/accent2026/data/saa

Requires openpyxl.
"""

import argparse
import csv
import random
import sys
from collections import Counter
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
DEFAULT_ARCHIVE = HERE.parents[1] / "Speech accent archive"
DEFAULT_SHEET = Path("excel_spreadsheets") / "speakers_textgrids_2026-10-07.xlsx"  # inside --archive-dir
COLUMNS = ["speaker", "native_language", "gender", "age", "country", "onset_age"]


def load_speakers(sheet_path):
    ws = openpyxl.load_workbook(sheet_path, read_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    speakers = []
    for row in rows:
        r = dict(zip(header, row))
        # Skip speakers without a TextGrid and rows that still need a human look
        if r.get("has_textgrid") != "yes" or r.get("review_flag"):
            continue
        # A few speakers have identical copies in two folders; use the first
        tg = r["textgrid_path"].split("; ")[0]
        r["speaker"] = Path(tg).stem
        r["textgrid_path"] = tg
        r["wav_path"] = str(Path(tg).with_suffix(".wav").as_posix())
        r["native_language"] = str(r["native_language"]).strip().lower()
        speakers.append(r)
    return speakers


def pick(speakers, n, seed):
    rng = random.Random(seed)
    l1 = sorted((s for s in speakers if s["native_language"] == "english"),
                key=lambda s: s["speaker"])
    l2 = sorted((s for s in speakers if s["native_language"] != "english"),
                key=lambda s: s["speaker"])

    # L1: half female, half male
    chosen_l1 = []
    for i, gender in enumerate(["female", "male"]):
        pool = [s for s in l1 if s["gender"] == gender]
        k = n // 2 + (n % 2 if i == 0 else 0)
        chosen_l1 += rng.sample(pool, k)

    # L2: one speaker from each of the n most common languages, alternating gender
    langs = [lang for lang, _ in Counter(s["native_language"] for s in l2).most_common(n)]
    chosen_l2 = []
    for i, lang in enumerate(langs):
        pool = [s for s in l2 if s["native_language"] == lang]
        want = "female" if i % 2 == 0 else "male"
        pool = [s for s in pool if s["gender"] == want] or pool
        chosen_l2.append(rng.choice(pool))

    return [("L1", s) for s in chosen_l1] + [("L2", s) for s in chosen_l2]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--n", type=int, default=10, help="speakers per group (default 10)")
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--name", default=None, help="output name (default pilot_<2n>)")
    ap.add_argument("--archive-dir", type=Path, default=DEFAULT_ARCHIVE,
                    help="Speech Accent Archive folder (default: the one in this project)")
    ap.add_argument("--spreadsheet", type=Path, default=None,
                    help=f"speaker spreadsheet (default: <archive-dir>/{DEFAULT_SHEET.as_posix()})")
    ap.add_argument("--out-dir", type=Path, default=HERE / "subsets")
    args = ap.parse_args()

    sheet = args.spreadsheet or args.archive_dir / DEFAULT_SHEET
    if not sheet.exists():
        sys.exit(f"Spreadsheet not found: {sheet}\nCopy excel_spreadsheets/ from Box first, or pass --spreadsheet.")

    name = args.name or f"pilot_{2 * args.n}"
    chosen = pick(load_speakers(sheet), args.n, args.seed)

    # Before the audio is copied (e.g. on Savio) the files won't be there yet; that's fine
    paths = [p for _, s in chosen for p in (s["textgrid_path"], s["wav_path"])]
    missing = [p for p in paths if not (args.archive_dir / p).exists()]
    if missing:
        print(f"note: {len(missing)} of {len(paths)} files are not in {args.archive_dir} yet "
              "(expected if you haven't copied them).\n", file=sys.stderr)

    args.out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = args.out_dir / f"{name}_speakers.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["group"] + COLUMNS + ["textgrid_path", "wav_path"])
        for group, s in chosen:
            w.writerow([group] + [s.get(c) for c in COLUMNS] + [s["textgrid_path"], s["wav_path"]])

    list_path = args.out_dir / f"{name}_files.txt"
    with open(list_path, "w", newline="\n", encoding="utf-8") as f:
        for _, s in chosen:
            f.write(s["textgrid_path"] + "\n" + s["wav_path"] + "\n")

    for group, s in chosen:
        print(f"{group}  {s['speaker']:<16} {s['native_language']:<12} {s['gender']:<7} {s['country']}")
    print(f"\nwrote {csv_path}\nwrote {list_path}")


if __name__ == "__main__":
    main()
