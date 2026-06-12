"""Organize messy HED files into HoleID/ATV and HoleID/OTV.

By default this script prints a dry-run plan. Use --apply to copy files into the
organized folder structure, or add --mode move when you really want to relocate
the originals.
"""

from __future__ import annotations

import argparse
import csv
import shutil
from dataclasses import dataclass
from pathlib import Path


DEFAULT_SOURCE_DIR = Path(r"C:\2026_RTIO\West Angelas\Sent")
PROJECT_ROOT = Path(__file__).resolve().parents[1]

OTV_MARKERS = ("OPTV", "OTV", "OBI")
ATV_MARKERS = ("BHTV", "ATV", "ABI", "BTV", "ATS")


@dataclass(frozen=True)
class OrganizeAction:
    source: Path
    destination: Path
    hole: str
    data_type: str


def infer_hole_id(path: Path) -> str | None:
    stem = path.stem
    if "_" not in stem:
        return None
    return stem.split("_", 1)[0]


def contains_marker(path: Path, markers: tuple[str, ...]) -> bool:
    text = " ".join((path.stem, *path.parts)).upper()
    return any(marker in text for marker in markers)


def classify_file(path: Path) -> str | None:
    if path.suffix.lower() != ".hed":
        return None
    if "LGX" in (part.upper() for part in path.parts) or "LGX" in path.stem.upper():
        return None
    if contains_marker(path, OTV_MARKERS):
        return "OTV"
    if contains_marker(path, ATV_MARKERS):
        return "ATV"
    return None


def unique_destination(destination: Path) -> Path:
    if not destination.exists():
        return destination

    counter = 2
    while True:
        candidate = destination.with_name(f"{destination.stem}_{counter}{destination.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def build_plan(source_dir: Path, output_dir: Path) -> list[OrganizeAction]:
    actions: list[OrganizeAction] = []
    reserved_destinations: set[Path] = set()
    for source in sorted(source_dir.rglob("*")):
        if not source.is_file():
            continue

        try:
            source.relative_to(output_dir)
            continue
        except ValueError:
            pass

        hole = infer_hole_id(source)
        data_type = classify_file(source)
        if not hole or not data_type:
            continue

        current_type_dir = output_dir / hole / data_type
        try:
            source.relative_to(current_type_dir)
            continue
        except ValueError:
            pass

        destination = output_dir / hole / data_type / source.name
        if destination.exists():
            continue
        while destination in reserved_destinations:
            destination = destination.with_name(f"{destination.stem}_copy{destination.suffix}")
        reserved_destinations.add(destination)
        if source.resolve() == destination.resolve():
            continue
        actions.append(
            OrganizeAction(
                source=source,
                destination=unique_destination(destination),
                hole=hole,
                data_type=data_type,
            )
        )
    return actions


def write_manifest(actions: list[OrganizeAction], manifest_path: Path) -> None:
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with manifest_path.open("w", newline="", encoding="utf-8") as manifest_file:
        writer = csv.DictWriter(
            manifest_file,
            fieldnames=["hole", "data_type", "source", "destination"],
        )
        writer.writeheader()
        for action in actions:
            writer.writerow(
                {
                    "hole": action.hole,
                    "data_type": action.data_type,
                    "source": action.source,
                    "destination": action.destination,
                }
            )


def apply_plan(actions: list[OrganizeAction], mode: str) -> None:
    for action in actions:
        action.destination.parent.mkdir(parents=True, exist_ok=True)
        if mode == "move":
            shutil.move(str(action.source), str(action.destination))
        else:
            shutil.copy2(action.source, action.destination)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE_DIR)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument(
        "--manifest",
        type=Path,
        default=PROJECT_ROOT / "logs" / "organize_hole_files_plan.csv",
    )
    parser.add_argument("--apply", action="store_true", help="Perform the copy/move. Default is dry-run.")
    parser.add_argument("--mode", choices=("copy", "move"), default="copy")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source_dir = args.source.expanduser().resolve()
    output_dir = (args.output or source_dir).expanduser().resolve()
    manifest_path = args.manifest.expanduser()
    if not manifest_path.is_absolute():
        manifest_path = PROJECT_ROOT / manifest_path
    manifest_path = manifest_path.resolve()

    if not source_dir.exists() or not source_dir.is_dir():
        raise SystemExit(f"Source folder not found: {source_dir}")

    actions = build_plan(source_dir, output_dir)
    print(f"Planned {len(actions)} file actions.")
    try:
        write_manifest(actions, manifest_path)
        print(f"Manifest: {manifest_path}")
    except PermissionError as exc:
        print(f"Manifest not written: {exc}")
    for action in actions[:20]:
        print(f"{action.data_type:3} {action.source} -> {action.destination}")
    if len(actions) > 20:
        print(f"... {len(actions) - 20} more in manifest")

    if args.apply:
        apply_plan(actions, args.mode)
        print(f"Applied plan using mode={args.mode}.")
    else:
        print("Dry-run only. Re-run with --apply to copy, or --apply --mode move to move.")


if __name__ == "__main__":
    main()
