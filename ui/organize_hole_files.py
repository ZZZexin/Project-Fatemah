"""Organize raw televiewer delivery files into HoleID / OTV | ATV | GPX.

CLI dry-run (default):
    python ui/organize_hole_files.py --source <path>

CLI apply:
    python ui/organize_hole_files.py --source <path> --apply [--mode move]
"""

from __future__ import annotations

import argparse
import configparser
import csv
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE_DIR = Path(r"C:\2026_RTIO\West Angelas\Sent")

OTV_MARKERS = ("OPTV", "OTV", "OBI")
ATV_MARKERS = ("BHTV", "ATV", "ABI", "BTV", "ATS")
REPORT_EXTENSIONS = {".pdf", ".doc", ".docx"}
SORTABLE_EXTENSIONS = {".hed", ".lox", ".lgx", ".las", ".xlsx"} | REPORT_EXTENSIONS


@dataclass(frozen=True)
class OrganizeAction:
    source: Path
    destination: Path
    hole: str
    data_type: str


# ---------------------------------------------------------------------------
# Hole-ID and file-type inference
# ---------------------------------------------------------------------------

def read_hed_wellname(hed_path: Path) -> str | None:
    """Read Wellname from a .hed INI file. Returns None on any failure."""
    # Try configparser first (handles standard INI structure)
    for encoding in ("utf-8", "latin-1"):
        try:
            cfg = configparser.ConfigParser()
            cfg.read(str(hed_path), encoding=encoding)
            for section in ("WellInfo", "wellinfo", "WELLINFO"):
                if cfg.has_section(section):
                    val = cfg.get(section, "wellname", fallback=None)
                    if val and val.strip():
                        return val.strip()
        except Exception:
            pass

    # Fallback: scan lines for "Wellname=..."
    for encoding in ("utf-8", "latin-1"):
        try:
            for line in hed_path.read_text(encoding=encoding).splitlines():
                stripped = line.strip()
                if stripped.lower().startswith("wellname="):
                    val = stripped.split("=", 1)[1].strip()
                    if val:
                        return val
        except Exception:
            pass

    return None


def infer_hole_id(path: Path) -> str | None:
    """Return hole ID from filename prefix (most reliable).
    Falls back to reading [WellInfo]Wellname only when the filename has no underscore.
    """
    stem = path.stem
    if "_" in stem:
        return stem.split("_", 1)[0]
    # No underscore — try reading from .hed metadata
    if path.suffix.lower() == ".hed":
        name = read_hed_wellname(path)
        if name:
            return name
    return stem or None


def has_marker(path: Path, markers: tuple[str, ...]) -> bool:
    text = " ".join([path.stem] + list(path.parts)).upper()
    return any(m in text for m in markers)


def classify_file(path: Path) -> str | None:
    """Return 'OTV', 'ATV', 'GPX', 'Reports', or None (skip this file)."""
    ext = path.suffix.lower()

    # Report files: PDF / Word docs — only if clearly named with a hole prefix
    if ext in REPORT_EXTENSIONS:
        return "Reports" if "_" in path.stem else None

    # Drop LGX-derived .hed headers
    if ext == ".hed":
        parts_upper = [p.upper() for p in path.parts]
        if "LGX" in path.stem.upper() or "LGX" in parts_upper:
            return None

    if ext not in SORTABLE_EXTENSIONS:
        return None

    # Instrument type takes priority over GPX
    if has_marker(path, ATV_MARKERS):
        return "ATV"
    if has_marker(path, OTV_MARKERS):
        return "OTV"

    # .las / .xlsx without instrument markers → survey/GPX data
    if ext in (".las", ".xlsx"):
        return "GPX"

    return None


# ---------------------------------------------------------------------------
# ZIP extraction
# ---------------------------------------------------------------------------

def extract_zips(source_dir: Path) -> list[Path]:
    """Extract all .zip files anywhere under source_dir, including zips inside zips."""
    extracted: list[Path] = []
    seen: set[Path] = set()
    while True:
        new_zips = [z for z in sorted(source_dir.rglob("*.zip")) if z not in seen]
        if not new_zips:
            break
        for zip_path in new_zips:
            seen.add(zip_path)
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    for member in zf.namelist():
                        dest = source_dir / member
                        if not dest.exists():
                            zf.extract(member, source_dir)
                            extracted.append(source_dir / member)
            except (zipfile.BadZipFile, OSError):
                pass
    return extracted


# ---------------------------------------------------------------------------
# Plan builder
# ---------------------------------------------------------------------------

def build_plan(
    source_dir: Path,
    output_dir: Path,
    extract_zip: bool = False,
) -> list[OrganizeAction]:
    if extract_zip:
        extract_zips(source_dir)

    actions: list[OrganizeAction] = []
    reserved: set[Path] = set()

    for source in sorted(source_dir.rglob("*")):
        if not source.is_file():
            continue

        # Skip anything already inside output_dir when output ≠ source
        if output_dir != source_dir:
            try:
                source.relative_to(output_dir)
                continue
            except ValueError:
                pass

        data_type = classify_file(source)
        if not data_type:
            continue

        hole = infer_hole_id(source)
        if not hole:
            continue

        dest_dir = output_dir / hole / data_type

        # Skip files already anywhere inside the correct HoleID/DataType tree
        try:
            source.relative_to(dest_dir)
            continue
        except ValueError:
            pass

        destination = dest_dir / source.name

        if source.resolve() == destination.resolve():
            continue
        if destination.exists():
            continue

        while destination in reserved:
            destination = destination.with_name(f"{destination.stem}_copy{destination.suffix}")
        reserved.add(destination)

        actions.append(OrganizeAction(
            source=source,
            destination=destination,
            hole=hole,
            data_type=data_type,
        ))

    return actions


# ---------------------------------------------------------------------------
# Execution
# ---------------------------------------------------------------------------

def apply_plan(actions: list[OrganizeAction], mode: str = "copy") -> None:
    for action in actions:
        action.destination.parent.mkdir(parents=True, exist_ok=True)
        if mode == "move":
            shutil.move(str(action.source), str(action.destination))
        else:
            shutil.copy2(action.source, action.destination)


def write_manifest(actions: list[OrganizeAction], manifest_path: Path) -> None:
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with manifest_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["hole", "data_type", "source", "destination"])
        writer.writeheader()
        for a in actions:
            writer.writerow({
                "hole": a.hole, "data_type": a.data_type,
                "source": str(a.source), "destination": str(a.destination),
            })


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE_DIR)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument(
        "--manifest", type=Path,
        default=PROJECT_ROOT / "logs" / "organize_plan.csv",
    )
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--mode", choices=("copy", "move"), default="copy")
    parser.add_argument("--extract-zip", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    source_dir = args.source.expanduser().resolve()
    output_dir = (args.output or source_dir).expanduser().resolve()

    if not source_dir.is_dir():
        raise SystemExit(f"Source not found: {source_dir}")

    actions = build_plan(source_dir, output_dir, extract_zip=args.extract_zip)
    print(f"Planned {len(actions)} file actions.")

    try:
        write_manifest(actions, args.manifest.resolve())
        print(f"Manifest: {args.manifest.resolve()}")
    except PermissionError as exc:
        print(f"Manifest not written: {exc}")

    for a in actions[:30]:
        print(f"  [{a.data_type:3}] {a.source.name}  ->  {a.destination}")
    if len(actions) > 30:
        print(f"  … {len(actions) - 30} more — see manifest")

    if args.apply:
        apply_plan(actions, mode=args.mode)
        print(f"Applied ({args.mode}).")
    else:
        print("Dry-run. Use --apply to execute.")


if __name__ == "__main__":
    main()
