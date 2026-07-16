#!/usr/bin/env python3
"""Select the explicit Python runtime shared by PDF provider scripts."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import sys


PROVIDER_PYTHON_ENV = "OPEN_OFFICE_PDF_PROVIDER_PYTHON"
REEXEC_GUARD_ENV = "_OPEN_OFFICE_PDF_PROVIDER_REEXEC_TARGET"


def configured_provider_python() -> Path | None:
    value = os.environ.get(PROVIDER_PYTHON_ENV, "").strip()
    if not value:
        return None
    candidate = Path(value).expanduser()
    if candidate.parent == Path("."):
        located = shutil.which(value)
        candidate = Path(located) if located else candidate
    candidate = candidate.resolve()
    if not candidate.is_file() or not os.access(candidate, os.X_OK):
        print(f"{PROVIDER_PYTHON_ENV} is not an executable file: {candidate}", file=sys.stderr)
        raise SystemExit(2)
    return candidate


def reexec_configured_provider_python() -> None:
    """Re-exec this script through the configured provider interpreter."""
    target = configured_provider_python()
    if target is None or Path(sys.executable).resolve() == target:
        return
    target_text = str(target)
    if os.environ.get(REEXEC_GUARD_ENV) == target_text:
        print(f"refusing a {PROVIDER_PYTHON_ENV} re-exec loop through {target_text}", file=sys.stderr)
        raise SystemExit(2)
    script = Path(sys.argv[0]).expanduser()
    if not script.is_absolute():
        script = (Path.cwd() / script).resolve()
    environment = dict(os.environ)
    environment[REEXEC_GUARD_ENV] = target_text
    os.execve(target_text, [target_text, str(script), *sys.argv[1:]], environment)
