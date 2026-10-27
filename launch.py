#!/usr/bin/env python3
"""Repository-aware launcher for TL1 Generator.

This launcher is intentionally GitHub-Workflow-free. It runs directly from any
clone, fork, or downloaded ZIP of the repository.

Behavior:
- Finds the repository root from the location of this file.
- Finds Python 3.10+ on Windows/macOS/Linux.
- Creates/reuses .venv in the repository.
- Installs requirements only when requirements.txt changes or --reinstall is used.
- Launches the GUI by default, or the CLI engine with --cli.
- Optionally performs a safe `git pull --ff-only` with --update when the checkout
  is clean. No GitHub Actions or workflow files are required.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional

APP_TITLE = "TL1 Generator"
MIN_PYTHON = (3, 10)
VENV_DIRNAME = ".venv"
STAMP_RELATIVE = Path(".tools_tl1") / "requirements.sha256"


def capture(cmd: List[str], cwd: Path) -> Optional[str]:
    try:
        result = subprocess.run(
            cmd,
            cwd=str(cwd),
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
        )
        return result.stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return None


def repository_root() -> Path:
    """Return the repository/project root without depending on its name."""
    here = Path(__file__).resolve().parent

    # Prefer the folder containing this launcher and the package. This also works
    # for GitHub Download ZIP, where .git metadata does not exist.
    if (here / "tl1_generator" / "__main__.py").is_file():
        return here

    # Defensive fallback: walk upward in case launch.py is placed in a subfolder.
    for parent in here.parents:
        if (parent / "tl1_generator" / "__main__.py").is_file():
            return parent

    raise RuntimeError(
        "Could not locate the repository root. Keep launch.py in the repository "
        "that contains the tl1_generator folder."
    )


def _python_candidates() -> List[List[str]]:
    if os.name == "nt":
        return [
            ["py", "-3.13"],
            ["py", "-3.12"],
            ["py", "-3.11"],
            ["py", "-3.10"],
            ["py", "-3"],
            ["python"],
        ]
    return [
        ["python3.13"],
        ["python3.12"],
        ["python3.11"],
        ["python3.10"],
        ["python3"],
        ["python"],
    ]


def bootstrap_python() -> List[str]:
    """Find Python 3.10+ and return its command prefix."""
    for candidate in _python_candidates():
        exe = shutil.which(candidate[0])
        if not exe:
            continue
        cmd = [exe, *candidate[1:]]
        probe = subprocess.run(
            [*cmd, "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if probe.returncode == 0:
            return cmd

    raise RuntimeError(
        "Python 3.10 or newer was not found. Install Python 3.10+ and run the launcher again."
    )


def venv_python(root: Path) -> Path:
    if os.name == "nt":
        return root / VENV_DIRNAME / "Scripts" / "python.exe"
    return root / VENV_DIRNAME / "bin" / "python"


def requirements_fingerprint(root: Path) -> str:
    req = root / "requirements.txt"
    digest = hashlib.sha256()
    digest.update(req.read_bytes() if req.exists() else b"")
    return digest.hexdigest()


def ensure_environment(root: Path, reinstall: bool = False) -> Path:
    """Create/reuse .venv and install dependencies only when needed."""
    py = venv_python(root)
    if not py.exists():
        host = bootstrap_python()
        print(f"[{APP_TITLE}] Creating local environment: {root / VENV_DIRNAME}")
        subprocess.run([*host, "-m", "venv", str(root / VENV_DIRNAME)], cwd=str(root), check=True)

    stamp = root / STAMP_RELATIVE
    fingerprint = requirements_fingerprint(root)
    old_fingerprint = stamp.read_text(encoding="utf-8").strip() if stamp.exists() else ""

    if reinstall or old_fingerprint != fingerprint:
        requirements = root / "requirements.txt"
        print(f"[{APP_TITLE}] Installing/updating dependencies...")
        subprocess.run([str(py), "-m", "pip", "install", "--upgrade", "pip"], cwd=str(root), check=True)
        if requirements.exists():
            subprocess.run(
                [str(py), "-m", "pip", "install", "-r", str(requirements)],
                cwd=str(root),
                check=True,
            )
        stamp.parent.mkdir(parents=True, exist_ok=True)
        stamp.write_text(fingerprint + "\n", encoding="utf-8")

    return py


def safe_update(root: Path) -> None:
    """Optionally fast-forward a clean Git clone. Never overwrite local edits."""
    git = shutil.which("git")
    if not git or not (root / ".git").exists():
        print(f"[{APP_TITLE}] Git clone not detected; skipping update.")
        return

    status = capture([git, "status", "--porcelain"], root)
    if status:
        print(f"[{APP_TITLE}] Local changes detected; skipping git pull.")
        return

    print(f"[{APP_TITLE}] Updating clone with git pull --ff-only...")
    result = subprocess.run([git, "pull", "--ff-only"], cwd=str(root))
    if result.returncode != 0:
        print(f"[{APP_TITLE}] Update was not completed. Launching the local copy instead.")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Launch TL1 Generator directly from a GitHub clone/fork or downloaded ZIP."
    )
    parser.add_argument("--cli", action="store_true", help="Launch the terminal CLI instead of the GUI.")
    parser.add_argument("--update", action="store_true", help="Safely git pull --ff-only before launch.")
    parser.add_argument("--reinstall", action="store_true", help="Force dependency reinstall.")
    parser.add_argument("--setup-only", action="store_true", help="Create .venv/install dependencies, then exit.")
    args = parser.parse_args()

    try:
        root = repository_root()
        print(f"[{APP_TITLE}] Repository root: {root}")

        if args.update:
            safe_update(root)

        py = ensure_environment(root, reinstall=args.reinstall)

        if args.setup_only:
            print(f"[{APP_TITLE}] Environment is ready.")
            return 0

        module = "tl1_generator.engine" if args.cli else "tl1_generator"
        print(f"[{APP_TITLE}] Launching {'CLI' if args.cli else 'desktop application'}...")
        return subprocess.call([str(py), "-m", module], cwd=str(root))

    except (RuntimeError, OSError, subprocess.CalledProcessError) as exc:
        print(f"[{APP_TITLE}] ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
