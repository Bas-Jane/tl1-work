# TL1 Generator

**Launch-only GitHub repository package**  
**Desktop app:** v1.3.7  
**TL1 engine:** v2.118.0  
**Platforms:** Windows 10/11, macOS, Linux

TL1 Generator converts supported CIQ Excel workbooks into TL1 configuration files. This repository is intentionally designed to run **directly from a GitHub clone, fork, or downloaded ZIP**.

## Repository contents

```text
TL1_Generator/
├── tl1_generator/
│   ├── __init__.py
│   ├── __main__.py
│   ├── engine.py
│   └── gui.py
├── launch.py
├── launch_windows.bat
├── launch_macos.command
├── launch_linux.sh
├── requirements.txt
├── .gitignore
└── README.md
```

That is the complete launch repository. No `.github/workflows`, build scripts, PyInstaller files, or installer scripts are required.

## Requirements

- Python **3.10 or newer**
- Internet access on the **first launch** so `openpyxl` can be installed into the repository-local `.venv`
- Tk support for the desktop GUI

On standard Windows and macOS Python installations, Tk is normally included. On Debian/Ubuntu Linux you may need:

```bash
sudo apt install python3-tk python3-venv
```

## Launch from GitHub

### 1. Clone the repository

```bash
git clone <your-repository-url>
cd <your-repository-folder>
```

The launcher does **not** depend on the repository owner, repository name, branch, or local path. Forks and renamed clones work without editing the scripts.

You can also use **GitHub → Code → Download ZIP**, extract it, and run the same launcher. Git is not required unless you want to use `--update`.

### 2. Windows

Double-click:

```text
launch_windows.bat
```

Or run:

```powershell
launch_windows.bat
```

### 3. macOS

First time only:

```bash
chmod +x launch_macos.command
```

Then:

```bash
./launch_macos.command
```

### 4. Linux

First time only:

```bash
chmod +x launch_linux.sh
```

Then:

```bash
./launch_linux.sh
```

## What `launch.py` does

On the first launch it automatically:

1. Locates the repository from the launcher file itself.
2. Finds Python 3.10+.
3. Creates a repository-local `.venv`.
4. Installs dependencies from `requirements.txt`.
5. Starts the TL1 Generator desktop GUI.

On later launches it reuses `.venv`. Dependencies are reinstalled only if `requirements.txt` changes or `--reinstall` is supplied.

## Direct launcher commands

Launch the desktop GUI:

```bash
python launch.py
```

Launch the terminal CLI:

```bash
python launch.py --cli
```

Prepare the environment without opening the application:

```bash
python launch.py --setup-only
```

Force dependency reinstall:

```bash
python launch.py --reinstall
```

Safely update a clean Git clone and then launch:

```bash
python launch.py --update
```

`--update` performs:

```text
git pull --ff-only
```

If local changes are present, the launcher skips the pull rather than overwriting them.

## Normal application workflow

1. Launch TL1 Generator.
2. Select the CIQ `.xlsx` or `.xlsm` workbook.
3. Load the workbook.
4. Select one or more sheets.
5. Scan for `CLI Command:` sections.
6. Select the sections to process.
7. Choose the output folder and filename stem.
8. Generate TL1.
9. Review `.tl1` output and `.warnings.txt` when generated.

The bundled engine supports the CIQ/TL1 processing rules developed for the tool, including target-value layouts, referred sheets, arrow/continuation rows, NE groups, timers, large-output splitting, case-preservation rules, skipped values, and large workbook handling.

## Updating the repository

Normal Git workflow:

```bash
git pull
python launch.py
```

Or let the launcher perform a safe fast-forward update:

```bash
python launch.py --update
```

## Files that should not be committed

`.gitignore` excludes the local virtual environment and generated TL1 files:

```text
.venv/
.tools_tl1/
__pycache__/
*.tl1
*.warnings.txt
```

## Troubleshooting

### Python not found

Install Python 3.10+ and ensure it is available as `py`, `python`, or `python3`.

### `No module named tkinter`

Linux:

```bash
sudo apt install python3-tk
```

### Virtual environment problem

Delete `.venv` and launch again, or run:

```bash
python launch.py --reinstall
```

### Repository has local changes

`python launch.py --update` intentionally refuses to pull over local modifications. Commit, stash, or discard the changes first, then run it again.

## Repository design

This package deliberately bypasses GitHub workflow automation. GitHub is used only as the source repository. Setup and launch happen on the user's own Windows, macOS, or Linux machine.

## Version synchronization

This GitHub launch package bundles the same TL1 processing engine as the standalone/local script.

- Standalone/local engine: `excel_to_tl1_generator_v2_103_0.py`
- GitHub package engine: `tl1_generator/engine.py`
- Bundled engine version: `2.118.0`

When the TL1 engine is updated, update both the standalone script and `tl1_generator/engine.py` together so local and GitHub launches produce the same output.


## Output folders

Each generation creates a new timestamped run folder under the selected output directory:

```text
<output-folder>/
└── YYYYMMDD_HHMMSS/
    ├── <stem>_v2_105_0.tl1
    ├── <stem>_v2_105_0_part002.tl1   # when splitting is required
    └── <stem>_v2_105_0.warnings.txt  # when warnings exist
```

This keeps separate runs isolated and prevents output from different executions from being mixed together.


## Current engine behavior

Bundled TL1 engine: `v2.118.0`

Recent parser improvements include:
- section-scoped `Out of Scope` detection so a following section cannot incorrectly suppress the current section;
- CLOG target rows stop before `CLOG Event Details` / `Details hidden` documentation tables;
- PM File Storage Period (`MEAS-KDAY`) generation from paired `*-day` / `*-hour` parameters when CLI Name cells are blank;
- `SVC="{all services}"` formatting for Service Overload Control;
- timestamped output run folders.


## Timestamped run folders

Every generation is written into a new execution-time folder created by the TL1 engine itself:

```text
<selected-output-folder>/
└── YYYYMMDD_HHMMSS/
    ├── <stem>_v2_107_0.tl1
    ├── <stem>_v2_107_0_part002.tl1
    └── <stem>_v2_107_0.warnings.txt
```

Both the standalone script and the GitHub GUI use the same engine-level output routing, so neither can bypass the timestamp folder.


## Engine v2.118.0 rules

- `CHG-OVLD-CNTL` rows with `SVC="{all services}"` are skipped.
- `CHG-VES-CNTL` rows with `EVENT={ALL}` are skipped.
- `CHG-MEAS-KDAY` removes the `RAW-` prefix from `TYPE`, for example `RAW-5MIN -> 5MIN`.
- TL1 files and warnings are written into the same timestamped execution folder.


## Engine v2.118.0 default-only rule

The generator skips a configuration record when every assigned **non-key**
parameter is expressed as a CIQ default wrapper such as:

```text
[default = 24000]
[default = 95]
[default = status-allow]
```

Explicit key fields such as service names, fault codes, or type identifiers do
not prevent the skip. This rule is generic and applies to all commands.

Intentional default-only skips are reported as `Skipped default-only ...`
rather than `No TL1 rows generated`.

The parser also stops at the next CIQ section header so status text such as
`Complete` cannot leak into the previous section's TL1 output.


## Engine v2.118.0 update

- `MEAS-KDAY` sections are intentionally skipped from TL1 generation.
- The skip is written to the warnings report as `Skipped MEAS-KDAY section ...`.
- TL1 output and the warnings report are generated in the same timestamped execution folder.


## Engine v2.118.0 CSDB update

The generator now supports the CSDB layout:

```text
CRTE-CSDB:LOCALIP,LOCALPORT,[REMOTEIP],[REMOTEPORT],NAME;
```

Rows with `csdb-info-single-redun-type=local` map to `LOCALIP/LOCALPORT`.
Rows with `csdb-info-single-redun-type=remote` map to `REMOTEIP/REMOTEPORT`.
Single-value braces around IP addresses are removed automatically.
