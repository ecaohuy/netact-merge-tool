"""Merge NetAct CR files (xlsx / csv) across site folders.

Rules:
- Files are grouped by (step number, NetAct tag, extension); NetAct01 and NetAct09
  files with the same number are NOT merged together.
- Steps are renumbered to the standard 7-step layout, based on the folder's highest step:
    6 steps (no "delete external EXENCE"): 1->1, 2->2, 3->3, 4->5, 5->6, 6->7
    8 steps (extra "delete LNADJGNB"):     1->1, 2->2, 3->3, 4->4.1, 5->4, 6->5, 7->6, 8->7
- Header (first 2 rows + following blank rows) is kept once; only values are merged.
- A group with a single source file is copied unchanged (all sheets kept).
- For xlsx, only the "CR detail new" sheet is merged. Side-by-side tables in that
  sheet (e.g. LNADJ | LNREL) are stacked independently so no gaps appear.

Input can be a folder or an archive (.rar / .zip); output can be a folder or an
archive (.rar / .zip). Extracting .rar needs `unrar` (or 7-Zip); creating .rar needs
`rar` (Ubuntu: sudo apt install rar unrar; Windows: WinRAR, or 7-Zip for extracting).

Several inputs can be given; all their site folders are merged together.

Usage: uv run --with openpyxl python merge.py input [input ...] [-o out_dir|out.rar|out.zip]
  default output: <input>/merged for one folder, <first input>_merged.rar otherwise
GUI:   uv run --with openpyxl python merge_gui.py
"""
import csv
import re
import shutil
import subprocess
import argparse
import os
import sys
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

import openpyxl

SHEET = "CR detail new"
HEADER_ROWS = 2
# highest step number in a folder -> {old step: new step}
RENUMBER = {
    6: {1: "1", 2: "2", 3: "3", 4: "5", 5: "6", 6: "7"},
    8: {1: "1", 2: "2", 3: "3", 4: "4.1", 5: "4", 6: "5", 7: "6", 8: "7"},
}
NAME_RE = re.compile(r"^(\d+)\.(.*?)(?:_(NetAct\d+))?$", re.IGNORECASE)


def parse(path: Path):
    m = NAME_RE.match(path.stem)
    if not m:
        return None
    return int(m.group(1)), m.group(2), (m.group(3) or "")


def input_files(folder: Path):
    return [
        f for f in sorted(folder.iterdir())
        if f.is_file() and not f.name.startswith((".~", "~$"))
        and f.suffix.lower() in (".xlsx", ".csv") and parse(f)
    ]


def site_folders(root: Path):
    """All folders under root (at any depth) holding input files, except merged/ output."""
    return sorted(
        d for d in [root, *root.rglob("*")]
        if d.is_dir() and "merged" not in (x.lower() for x in d.relative_to(root).parts) and input_files(d)
    )


def collect(roots):
    groups = defaultdict(list)  # (num, tag, ext) -> [(folder, path, stem_name)]
    for folder in (f for root in roots for f in site_folders(root)):
        files = input_files(folder)
        nums = {parse(f)[0] for f in files}
        remap = RENUMBER.get(max(nums), {}) if nums else {}
        for f in files:
            num, name, tag = parse(f)
            num = remap.get(num, str(num))
            groups[(num, tag, f.suffix.lower())].append((folder.name, f, name))
    return groups


def out_name(num, name, tag, ext):
    return f"{num}.{name}{'_' + tag if tag else ''}{ext}"


def merge_csv(items, out: Path):
    header, data = None, []
    for _, f, _ in items:
        with open(f, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        first_data = next(
            (i for i in range(HEADER_ROWS, len(rows)) if any(c.strip() for c in rows[i])),
            len(rows),
        )
        if header is None:
            header = rows[:first_data]
        elif rows[:HEADER_ROWS] != header[:HEADER_ROWS]:
            print(f"  WARNING: header differs in {f}")
        data += [r for r in rows[first_data:] if any(c.strip() for c in r)]
    with open(out, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh, lineterminator="\r\n").writerows(header + data)
    return len(data)


def blocks_of(ws):
    """Column ranges [start, end] of the side-by-side tables, from row 1 labels."""
    starts = [c.column for c in ws[1] if c.value not in (None, "")]
    ends = [s - 1 for s in starts[1:]] + [ws.max_column]
    return list(zip(starts, ends))


def merge_xlsx(items, out: Path):
    template = items[0][1]
    wb_out = openpyxl.load_workbook(template)
    for name in wb_out.sheetnames:
        if name != SHEET:
            del wb_out[name]
    ws_out = wb_out[SHEET]
    blocks = blocks_of(ws_out)
    header_vals = [[c.value for c in row] for row in ws_out.iter_rows(max_row=HEADER_ROWS)]

    # first data row in template (header + blank spacer rows are preserved)
    first_data = HEADER_ROWS + 1
    while first_data <= ws_out.max_row and all(
        c.value in (None, "") for c in ws_out[first_data]
    ):
        first_data += 1
    if first_data > ws_out.max_row:
        first_data = HEADER_ROWS + 3

    per_block = [[] for _ in blocks]
    for _, f, _ in items:
        ws = openpyxl.load_workbook(f, data_only=True)[SHEET]
        vals = [[c.value for c in row] for row in ws.iter_rows(max_row=HEADER_ROWS)]
        if vals != header_vals:
            print(f"  WARNING: header differs in {f}")
        for row in ws.iter_rows(min_row=HEADER_ROWS + 1, values_only=True):
            for i, (s, e) in enumerate(blocks):
                seg = list(row[s - 1:e])
                if any(v not in (None, "") for v in seg):
                    per_block[i].append(seg)

    # clear old data, then write stacked blocks
    ws_out.delete_rows(first_data, ws_out.max_row)
    for (s, _), rows in zip(blocks, per_block):
        for r, seg in enumerate(rows):
            for c, v in enumerate(seg):
                ws_out.cell(row=first_data + r, column=s + c, value=v)
    wb_out.save(out)
    return [len(b) for b in per_block]


ARCHIVES = (".rar", ".zip")


class MergeError(Exception):
    pass


def find_tool(*names):
    """Locate an external program. Besides PATH, looks in ~/.local/bin (not always on
    PATH when started from a desktop launcher) and the WinRAR / 7-Zip install folders
    on Windows."""
    dirs = [os.environ.get("PATH", ""), str(Path.home() / ".local/bin")]
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        if os.environ.get(env):
            dirs += [str(Path(os.environ[env]) / "WinRAR"), str(Path(os.environ[env]) / "7-Zip")]
    path = os.pathsep.join(dirs)
    for n in names:
        found = shutil.which(n, path=path)
        if found:
            return found
    return None


def run(cmd):
    # CREATE_NO_WINDOW: no console flashing up when run from the Windows .exe
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, creationflags=flags)


def extract(archive: Path, dest: Path):
    if archive.suffix.lower() == ".zip":
        zipfile.ZipFile(archive).extractall(dest)
        return
    dest.mkdir(parents=True, exist_ok=True)
    tool = find_tool("unrar", "rar")
    if tool:
        run([tool, "x", "-o+", "-inul", str(archive), str(dest) + os.sep])
    else:
        run([find_tool("7z", "7zz"), "x", "-y", f"-o{dest}", str(archive)])


def pack(src_dir: Path, archive: Path):
    archive.unlink(missing_ok=True)
    if archive.suffix.lower() == ".zip":
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(src_dir.rglob("*")):
                z.write(f, f.relative_to(src_dir.parent))
    else:
        # -ep1: store paths relative to src_dir's parent, so the archive holds <name>/...
        run([find_tool("rar"), "a", "-r", "-ep1", "-inul", str(archive), str(src_dir)])


def is_archive(p: Path):
    return p.is_file() and p.suffix.lower() in ARCHIVES


def default_output(sources):
    first = sources[0]
    if len(sources) == 1 and first.is_dir():
        return first / "merged"
    return first.with_name(first.stem + "_merged.rar")


def merge(sources, out: Path, log=print):
    """Merge all site folders found in `sources` (folders and/or .rar/.zip files) into
    `out` (a folder, or a .rar/.zip archive)."""
    sources = [Path(s).resolve() for s in sources]
    out = Path(out).resolve()
    for s in sources:
        if not s.exists():
            raise MergeError(f"Input not found: {s}")
    needs = []
    if any(s.suffix.lower() == ".rar" for s in sources) and not find_tool("unrar", "rar", "7z", "7zz"):
        needs.append("unrar (or WinRAR / 7-Zip)")
    if out.suffix.lower() == ".rar" and not find_tool("rar"):
        needs.append("rar")
    if needs:
        raise MergeError(f"{' and '.join(needs)} not installed. Linux: sudo apt install rar unrar; "
                         "Windows: install WinRAR or 7-Zip")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        roots = []
        for i, s in enumerate(sources):
            if is_archive(s):
                log(f"Extracting {s.name} ...")
                root = tmp / f"input{i}"
                extract(s, root)
                roots.append(root)
            else:
                roots.append(s)
        out_is_archive = out.suffix.lower() in ARCHIVES
        out_dir = tmp / out.stem if out_is_archive else out
        out_dir.mkdir(parents=True, exist_ok=True)
        count = merge_all(roots, out_dir, log)
        if count == 0:
            raise MergeError("No numbered .xlsx / .csv files found in the input")
        if out_is_archive:
            pack(out_dir, out)
    log(f"\nDone: {count} files -> {out}")
    return out


def merge_all(roots, out_dir: Path, log=print):
    groups = collect(roots)
    for (num, tag, ext) in sorted(groups, key=lambda k: (float(k[0]), k[1], k[2])):
        items = groups[(num, tag, ext)]
        # name the output after the file from a full (7-step) folder
        name = items[0][2]
        out = out_dir / out_name(num, name, tag, ext)
        log(f"{out.name}  <-  " + ", ".join(f"{d}/{f.name}" for d, f, _ in items))
        if len(items) == 1:
            # only one source file: nothing to merge, copy the original unchanged
            shutil.copy2(items[0][1], out)
            log("  copied (single source)")
            continue
        n = merge_csv(items, out) if ext == ".csv" else merge_xlsx(items, out)
        log(f"  rows: {n}")
    return len(groups)


def main():
    ap = argparse.ArgumentParser(description="Merge NetAct CR files across site folders.")
    ap.add_argument("inputs", nargs="*", default=["."], help="folders and/or .rar/.zip files")
    ap.add_argument("-o", "--output", help="output folder or .rar/.zip file")
    args = ap.parse_args()
    sources = [Path(s).resolve() for s in args.inputs]
    out = Path(args.output) if args.output else default_output(sources)
    try:
        merge(sources, out)
    except MergeError as e:
        sys.exit(f"ERROR: {e}")


if __name__ == "__main__":
    main()
