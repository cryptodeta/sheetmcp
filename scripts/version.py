#!/usr/bin/env python3
"""Release version helper. VERSION.md (first line) is the single source of truth.

python scripts/version.py get            print the version from VERSION.md
python scripts/version.py sync           write it into pyproject.toml, sheetmcp.__version__ and uv.lock
python scripts/version.py check          exit 1 if any file is out of sync with VERSION.md
python scripts/version.py release-check  compare with the latest v* git tag (key=value lines for $GITHUB_OUTPUT)
python scripts/version.py notes [ver]    print the CHANGELOG.md section for a version
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
from functools import cmp_to_key
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SEMVER = re.compile(r"^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?$")


def read_version() -> str:
    text = (ROOT / "VERSION.md").read_text(encoding="utf-8")
    first = next((line for line in text.splitlines() if line.strip()), "")
    v = re.sub(r"^(?:#+\s*)?(?:version\s*:?\s*)?v?", "", first.strip(), flags=re.I).strip()
    if not SEMVER.match(v):
        raise SystemExit(f'VERSION.md: first line must be a semantic version like 1.2.3, got "{first}"')
    return v


def compare(a: str, b: str) -> int:
    ma, mb = SEMVER.match(a), SEMVER.match(b)
    assert ma and mb
    for i in (1, 2, 3):
        d = int(ma[i]) - int(mb[i])
        if d:
            return (d > 0) - (d < 0)
    pa, pb = ma[4], mb[4]
    if not pa and not pb:
        return 0
    if not pa:
        return 1  # 1.0.0 > 1.0.0-rc.1
    if not pb:
        return -1
    xa, xb = pa.split("."), pb.split(".")
    for i in range(max(len(xa), len(xb))):
        if i >= len(xa):
            return -1
        if i >= len(xb):
            return 1
        na, nb = xa[i].isdigit(), xb[i].isdigit()
        if na and nb and int(xa[i]) != int(xb[i]):
            return 1 if int(xa[i]) > int(xb[i]) else -1
        if na != nb:
            return -1 if na else 1
        if xa[i] != xb[i]:
            return -1 if xa[i] < xb[i] else 1
    return 0


def plan(version: str) -> dict[Path, str]:
    changes: dict[Path, str] = {}
    f = ROOT / "pyproject.toml"
    old = f.read_text(encoding="utf-8")
    new = re.sub(r'(?m)^version = "[^"]*"', f'version = "{version}"', old, count=1)
    if new != old:
        changes[f] = new
    for f in [ROOT / "src" / "sheetmcp" / "__init__.py"]:
        old = f.read_text(encoding="utf-8")
        new = re.sub(r'(?m)^__version__ = "[^"]*"', f'__version__ = "{version}"', old)
        if new != old:
            changes[f] = new
    return changes


def latest_tag() -> str | None:
    try:
        out = subprocess.run(
            ["git", "tag", "--list", "v*"], cwd=ROOT, capture_output=True, text=True, check=True
        ).stdout
    except (OSError, subprocess.CalledProcessError):
        return None
    tags = [t.strip()[1:] for t in out.splitlines() if SEMVER.match(t.strip()[1:])]
    return max(tags, key=cmp_to_key(compare)) if tags else None


def changelog_section(version: str) -> str:
    f = ROOT / "CHANGELOG.md"
    if not f.exists():
        return ""
    lines = f.read_text(encoding="utf-8").splitlines()
    head = re.compile(rf"^##\s+\[?v?{re.escape(version)}\]?(\s|$)")
    start = next((i for i, line in enumerate(lines) if head.match(line)), None)
    if start is None:
        return ""
    end = next((i for i in range(start + 1, len(lines)) if lines[i].startswith("## ")), len(lines))
    return "\n".join(lines[start + 1 : end]).strip()


def main(argv: list[str]) -> int:
    cmd = argv[0] if argv else "get"
    version = read_version()
    if cmd == "get":
        print(version)
    elif cmd == "sync":
        changes = plan(version)
        for f, text in changes.items():
            f.write_text(text, encoding="utf-8", newline="\n")
        if changes and shutil.which("uv"):
            subprocess.run(["uv", "lock"], cwd=ROOT, check=True)
        names = "\n".join(f"  {f.relative_to(ROOT).as_posix()}" for f in changes)
        print(f"Synced {version} into:\n{names}" if changes else f"All files already at {version}")
    elif cmd == "check":
        changes = plan(version)
        if changes:
            print(f"Out of sync with VERSION.md ({version}): {', '.join(map(str, changes))}", file=sys.stderr)
            print("Run: python scripts/version.py sync", file=sys.stderr)
            return 1
        print(f"All package versions match VERSION.md ({version})")
    elif cmd == "release-check":
        latest = latest_tag()
        release = latest is None or compare(version, latest) > 0
        msg = (
            f"Releasing v{version} (latest tag: {'v' + latest if latest else 'none'})."
            if release
            else f"Skipping release: VERSION.md is {version}, which is not greater than the latest tag "
            f"v{latest}. Bump VERSION.md to publish."
        )
        print(f"version={version}\nlatest={latest or ''}\nrelease={str(release).lower()}")
        print(f"prerelease={str('-' in version).lower()}")
        print(msg, file=sys.stderr)
        summary = os.environ.get("GITHUB_STEP_SUMMARY")
        if summary:
            with open(summary, "a", encoding="utf-8") as fh:
                fh.write(f"### Release check\n\n{msg}\n")
    elif cmd == "notes":
        v = argv[1] if len(argv) > 1 else version
        print(changelog_section(v) or f"Release v{v}. See CHANGELOG.md for details.")
    else:
        print(f"Unknown command: {cmd}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
