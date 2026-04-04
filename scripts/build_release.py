"""Build a release zip for PdbToExcel.

Builds wheels for all private dependencies and pnfl-pdbtoexcel itself,
copies release files (bat, config, docs, license) into a staging folder,
and zips the result into dist/.

Run from the pnfl-pdbtoexcel project root:
    py -3.13 scripts/build_release.py
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
PNFL_ROOT = PROJECT_ROOT.parent

# Private dependency projects (order doesn't matter — pip resolves it).
DEPENDENCY_PROJECTS = [
    PNFL_ROOT / "fbpro98-play",
    PNFL_ROOT / "fbpro98-gameplan",
    PNFL_ROOT / "pnfl-playpool",
    PROJECT_ROOT,
]


def get_version() -> str:
    """Read the version from pyproject.toml."""
    pyproject = PROJECT_ROOT / "pyproject.toml"
    match = re.search(r'^version\s*=\s*"(.+?)"', pyproject.read_text(), re.MULTILINE)
    if not match:
        raise RuntimeError("Could not find version in pyproject.toml")
    return match.group(1)


def build_wheel(project_dir: Path, output_dir: Path) -> None:
    """Build a wheel for a project into the output directory."""
    print(f"Building wheel: {project_dir.name}")
    subprocess.run(
        [
            sys.executable, "-m", "pip", "wheel",
            "--no-deps", "--wheel-dir", str(output_dir), str(project_dir),
        ],
        check=True,
    )


def main() -> None:
    version = get_version()
    release_name = f"PdbToExcel-v{version}"

    dist_dir = PROJECT_ROOT / "dist"
    staging = dist_dir / release_name
    packages_dir = staging / "packages"

    # Clean previous build.
    if staging.exists():
        shutil.rmtree(staging)
    packages_dir.mkdir(parents=True)

    # Build wheels.
    for project in DEPENDENCY_PROJECTS:
        build_wheel(project, packages_dir)

    # Copy release files (bat, ini, README, etc.).
    release_dir = PROJECT_ROOT / "release"
    for release_file in release_dir.glob("*"):
        if release_file.is_file():
            shutil.copy2(release_file, staging)

    # License.
    license_src = PROJECT_ROOT / "LICENSE.txt"
    if license_src.exists():
        shutil.copy2(license_src, staging)

    # Zip.
    zip_path = dist_dir / release_name
    shutil.make_archive(str(zip_path), "zip", dist_dir, release_name)

    print()
    print(f"Release built: {zip_path}.zip")


if __name__ == "__main__":
    main()
