#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Run pyexam to build LaTeX/PDF from exported YAML")
    p.add_argument("input_yaml", help="pyexam YAML file")
    p.add_argument("output_dir", help="Output directory for pyexam artifacts")
    p.add_argument("--format", choices=["latex", "pdf"], default="pdf", help="Output format (default: pdf)")
    p.add_argument("--solution", action="store_true", help="Build solution version (-s)")
    p.add_argument("--template-dir", default=None, help="Optional pyexam template directory")
    p.add_argument(
        "--python-module-fallback",
        action="store_true",
        help="If 'pyexam' CLI is not found, try 'python -m pyexam'",
    )
    p.add_argument("--dry-run", action="store_true", help="Print command only")
    return p


def main() -> int:
    args = build_parser().parse_args()
    input_yaml = Path(args.input_yaml)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if shutil.which("pyexam"):
        cmd = ["pyexam"]
    elif args.python_module_fallback:
        cmd = ["python", "-m", "pyexam"]
    else:
        print("ERROR: 'pyexam' CLI not found on PATH. Install pyexam or use --python-module-fallback.")
        return 1

    cmd += ["-i", str(input_yaml), "-f", args.format, "-o", str(output_dir)]
    if args.template_dir:
        cmd += ["-t", str(Path(args.template_dir))]
    if args.solution:
        cmd += ["-s"]

    print("Running:")
    print(" ".join(cmd))
    if args.dry_run:
        return 0

    proc = subprocess.run(cmd)
    return int(proc.returncode)


if __name__ == "__main__":
    raise SystemExit(main())
