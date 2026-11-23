0.2.0

# Release version

The first line of this file is the release version of every package in this repository
(`sheetmcp-core`, `sheetmcp`, `sheetmcp-cli`). It is the only thing you change to ship a release.

## How to release

1. Add a `## [x.y.z] - YYYY-MM-DD` section to `CHANGELOG.md` (it becomes the GitHub Release notes).
2. Replace the first line of this file with the new version, for example `0.2.0` -> `0.3.0`.
   Use semantic versioning; pre-releases such as `1.0.0-rc.1` are marked as pre-releases on GitHub.
3. Commit and push (or merge) to `main`.

The `Release` workflow then compares this version with the latest `v*` git tag:

- **Greater**: syncs the version into every `pyproject.toml`, `__version__` and `uv.lock`
  (`python scripts/version.py sync`), runs lint and tests, builds wheels and sdists, creates the tag `vX.Y.Z`
  and a GitHub Release with the files attached, then publishes to PyPI and to GitHub Packages as the Docker
  image `ghcr.io/cryptodeta/sheetmcp`.
- **Equal or lower**: nothing is tagged or published; the workflow finishes green with a skip note.

Run `python scripts/version.py sync` locally if you want the package files updated in the same commit.
