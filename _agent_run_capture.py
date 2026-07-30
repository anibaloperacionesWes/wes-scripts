"""One-off runner: execute control_nocturno and write exit code + file paths."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "_agent_run_capture_result.txt"
REPORTS = ROOT / "reports" / "control_nocturno"

def main() -> int:
    before = {p.resolve() for p in REPORTS.glob("*20260526*")} if REPORTS.is_dir() else set()
    proc = subprocess.run(
        [sys.executable, str(ROOT / "control_nocturno.py"), "--desde", "2026-05-26", "--hasta", "2026-05-26"],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
    )
    after = {p.resolve() for p in REPORTS.glob("*20260526*")} if REPORTS.is_dir() else set()
    created = sorted(after - before)
    lines = [
        f"EXIT_CODE={proc.returncode}",
        "STDOUT:",
        proc.stdout or "",
        "STDERR:",
        proc.stderr or "",
        "CREATED_FILES:",
    ]
    lines.extend(str(p) for p in created)
    if not created:
        lines.append("(none new; listing all *20260526* matches)")
        lines.extend(str(p.resolve()) for p in sorted(REPORTS.glob("*20260526*")))
    OUT.write_text("\n".join(lines), encoding="utf-8")
    return proc.returncode

if __name__ == "__main__":
    raise SystemExit(main())
