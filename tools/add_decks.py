#!/usr/bin/env python3
"""Put lecture PowerPoints on the site so students download them directly.

    python3 tools/add_decks.py PATH [PATH ...]

Each PATH is a .pptx, a folder, or a .zip (e.g. the "Download" zips Google Drive makes from the
NEUROETHOLGY-UPDATED LECTURES folder). The deck for each lecture is the file named in
source/lectures.json; when that name is absent, a single file whose name contains LectureN is used.
Chosen decks are copied to docs/files/lectures/NEUR411_LectureNN.pptx, the manifest is updated and
the site is rebuilt.
"""
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "source" / "lectures.json"
DEST = ROOT / "docs" / "files" / "lectures"
GITHUB_LIMIT = 100 * 1024 * 1024


def candidates(paths, scratch):
    """Yield (file name, Path) for every .pptx in the given files, folders and zips."""
    for p in map(Path, paths):
        if p.is_dir():
            for f in p.rglob("*.pptx"):
                yield f.name, f
        elif p.suffix.lower() == ".zip":
            with zipfile.ZipFile(p) as z:
                for info in z.infolist():
                    name = Path(info.filename).name
                    if name.lower().endswith(".pptx") and not name.startswith("._"):
                        out = scratch / f"{len(list(scratch.iterdir()))}_{name}"
                        with z.open(info) as src, open(out, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        yield name, out
        elif p.suffix.lower() == ".pptx":
            yield p.name, p


def main(paths):
    manifest = json.loads(MANIFEST.read_text())
    entries = {e["n"]: e for e in manifest["lectures"]}
    scratch = ROOT / ".add_decks_tmp"
    scratch.mkdir(exist_ok=True)
    try:
        found = {}
        for name, path in candidates(paths, scratch):
            m = re.search(r"Lecture\s*0*(\d+)", name, re.I)
            if m:
                found.setdefault(int(m.group(1)), []).append((name, path))
        DEST.mkdir(parents=True, exist_ok=True)
        added, unclear = [], []
        for n, entry in sorted(entries.items()):
            files = found.get(n, [])
            exact = [f for f in files if f[0] == entry.get("file")]
            pick = exact[0] if exact else (files[0] if len(files) == 1 else None)
            if pick is None:
                if files:
                    unclear.append(f"Lecture {n}: " + ", ".join(sorted(f[0] for f in files)))
                continue
            name, path = pick
            if path.stat().st_size > GITHUB_LIMIT:
                unclear.append(f"Lecture {n}: {name} is over GitHub's 100 MB file limit")
                continue
            local = f"NEUR411_Lecture{n:02d}.pptx"
            shutil.copyfile(path, DEST / local)
            entry.update(local=local, file=name, size=path.stat().st_size)
            added.append(n)
        MANIFEST.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")
    finally:
        shutil.rmtree(scratch, ignore_errors=True)

    print(f"added {len(added)} decks: {added}")
    missing = sorted(set(entries) - set(added))
    if unclear:
        print("needs a decision (set \"file\" in source/lectures.json and rerun):\n  " + "\n  ".join(unclear))
    if missing:
        print(f"no deck yet for lectures {missing}")
    subprocess.run([sys.executable, str(ROOT / "tools" / "build_site.py")], check=True)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1:])
