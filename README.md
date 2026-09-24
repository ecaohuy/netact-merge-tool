# NetAct CR Merge Tool

Merges NetAct CR files (`.xlsx` / `.csv`) from several site folders into one file per step.

## Rules
- Files are grouped by step number and NetAct tag. `NetAct01` and `NetAct09` files are never merged together.
- Steps are renumbered to the standard 7-step layout, based on the highest step in each folder:
  - 6 steps: 1→1, 2→2, 3→3, 4→5, 5→6, 6→7
  - 8 steps: 1→1, 2→2, 3→3, 4→4.1, 5→4, 6→5, 7→6, 8→7
- The header (2 rows plus the blank rows under it) is kept once, and only the values are merged.
- In `.xlsx` files only the `CR detail new` sheet is merged. Side-by-side tables (e.g. LNADJ | LNREL) are stacked separately.
- A file that exists in only one site folder is copied unchanged.
- `.~` lock files, `*_profile.xml` and any `merged/` folder are ignored.

## Usage
Windows: download `MergeTool.exe` from the Actions artifacts or Releases, then run it.
Select one or more `.rar` / `.zip` files, choose the output `.zip`, and press **Run**.
Reading `.rar` needs WinRAR or 7-Zip installed.

From source (needs [uv](https://docs.astral.sh/uv/)):
```
uv run merge_gui.py                                  # GUI
uv run merge.py a.rar b.rar -o merged.zip            # command line
uv run merge.py site_folders_dir                     # -> site_folders_dir/merged
```
On Linux, `.rar` support needs `rar` / `unrar` (`sudo apt install rar unrar`).
