#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import List

try:
    from kabuto_common import dump_json
except Exception:  # pragma: no cover
    import json
    def dump_json(obj, path):
        Path(path).write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding='utf-8')


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Run pyexam to build LaTeX/PDF from exported YAML")
    p.add_argument("input_path", help="pyexam YAML file OR directory containing exported YAML exam set")
    p.add_argument("output_dir", help="Output directory for pyexam artifacts")
    p.add_argument("--format", choices=["latex", "pdf"], default="pdf", help="Output format (default: pdf)")
    p.add_argument("--solution", action="store_true", help="Build solution version (-s)")
    p.add_argument("--template-dir", default=None, help="Optional pyexam template directory")
    p.add_argument(
        "--python-module-fallback",
        action="store_true",
        help="If 'pyexam' CLI is not found, try 'python -m pyexam'",
    )
    p.add_argument("--glob", default="*.yaml", help="Glob when input_path is a directory (default: *.yaml)")
    p.add_argument("--dry-run", action="store_true", help="Print command(s) only")
    return p


def _resolve_runner(python_module_fallback: bool) -> List[str] | None:
    if shutil.which("pyexam"):
        return ["pyexam"]
    if python_module_fallback:
        return ["python", "-m", "pyexam"]
    return None


def _build_cmd(base_cmd: List[str], input_yaml: Path, output_dir: Path, fmt: str, solution: bool, template_dir: str | None) -> List[str]:
    cmd = list(base_cmd)
    cmd += ["-i", str(input_yaml), "-f", fmt, "-o", str(output_dir)]
    if template_dir:
        cmd += ["-t", str(Path(template_dir))]
    if solution:
        cmd += ["-s"]
    return cmd


def _run_one(cmd: List[str], dry_run: bool) -> int:
    print("Running:")
    print(" ".join(cmd))
    if dry_run:
        return 0
    proc = subprocess.run(cmd)
    return int(proc.returncode)


def main() -> int:
    args = build_parser().parse_args()
    input_path = Path(args.input_path)
    output_root = Path(args.output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    base_cmd = _resolve_runner(args.python_module_fallback)
    if base_cmd is None:
        print("ERROR: 'pyexam' CLI not found on PATH. Install pyexam or use --python-module-fallback.")
        return 1

    created_at = datetime.now().isoformat(timespec='seconds')

    if input_path.is_file():
        cmd = _build_cmd(base_cmd, input_path, output_root, args.format, args.solution, args.template_dir)
        rc = _run_one(cmd, args.dry_run)
        manifest = {
            'created_at': created_at,
            'mode': 'single',
            'input_yaml': str(input_path),
            'output_dir': str(output_root),
            'format': args.format,
            'solution': bool(args.solution),
            'template_dir': args.template_dir,
            'dry_run': bool(args.dry_run),
            'command': cmd,
            'returncode': rc,
        }
        dump_json(manifest, output_root / 'pyexam_build_manifest.json')
        print(f"Build manifest -> {output_root / 'pyexam_build_manifest.json'}")
        return rc

    if not input_path.is_dir():
        print(f"ERROR: input_path does not exist: {input_path}")
        return 2

    yamls = sorted([p for p in input_path.glob(args.glob) if p.is_file()])
    # drop obvious non-exam YAML files if present (none expected, but be safe)
    yamls = [p for p in yamls if not p.name.endswith('.answer_key.yaml')]
    if not yamls:
        print(f"ERROR: no YAML files found in {input_path} matching {args.glob}")
        return 3

    records = []
    failures = 0
    for y in yamls:
        version_out = output_root / y.stem
        version_out.mkdir(parents=True, exist_ok=True)
        cmd = _build_cmd(base_cmd, y, version_out, args.format, args.solution, args.template_dir)
        rc = _run_one(cmd, args.dry_run)
        if rc != 0:
            failures += 1
        rec = {
            'input_yaml': str(y),
            'output_dir': str(version_out),
            'command': cmd,
            'returncode': rc,
        }
        records.append(rec)
        dump_json(rec, version_out / 'pyexam_build_manifest.json')

    batch_manifest = {
        'created_at': created_at,
        'mode': 'batch',
        'input_dir': str(input_path),
        'glob': args.glob,
        'output_dir': str(output_root),
        'format': args.format,
        'solution': bool(args.solution),
        'template_dir': args.template_dir,
        'dry_run': bool(args.dry_run),
        'count': len(records),
        'failures': failures,
        'records': records,
    }
    dump_json(batch_manifest, output_root / 'pyexam_batch_build_manifest.json')
    print(f"Batch build manifest -> {output_root / 'pyexam_batch_build_manifest.json'}")
    return 0 if failures == 0 else 4


if __name__ == "__main__":
    raise SystemExit(main())
