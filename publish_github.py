"""Publish the dashboard to GitHub Pages: https://ericliu76.github.io/market-dashboard/

Copies the static site (site/), the data files the page reads and the pipeline scripts into a local
clone of the public repo EricLiu76/market-dashboard, then commits and pushes; GitHub Pages serves
the repo root. Run it after build_data.py:

  python publish_github.py              commit and push
  python publish_github.py --no-push    copy and commit only (to preview first)
  python publish_github.py -m "text"    use your own commit message

The clone lives outside Google Drive (~/Repos/market-dashboard) so Drive never syncs .git; it is
cloned again if missing. Market_Data.xlsx and master_data_long.csv stay local: the page reads
master_data_file.csv, which holds the same data.
"""
import argparse
import shutil
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
REPO = Path.home() / "Repos" / "market-dashboard"
REMOTE = "https://github.com/EricLiu76/market-dashboard.git"
PAGES_URL = "https://ericliu76.github.io/market-dashboard/"
SGT = timezone(timedelta(hours=8))

# repo path -> local source
FILES = {
    "index.html": ROOT / "site" / "index.html",
    "dash-shim.js": ROOT / "site" / "dash-shim.js",
    "dashboard_page.html": ROOT / "site" / "dashboard_page.html",
    "README.md": ROOT / "site" / "README.md",
    "universe.csv": ROOT / "universe.csv",
    "build_data.py": ROOT / "build_data.py",
    "build_calendar.py": ROOT / "build_calendar.py",
    "publish_github.py": ROOT / "publish_github.py",
    "data/market_snapshot.csv": DATA / "market_snapshot.csv",
    "data/earnings_calendar.csv": DATA / "earnings_calendar.csv",
}
FILES.update({f"data/{p.name}": p for p in sorted(DATA.glob("master_data_file*.csv"))})


def git(*args, capture=False):
    r = subprocess.run(["git", "-C", str(REPO), *args], text=True, capture_output=capture)
    if r.returncode != 0:
        if capture:
            sys.stderr.write(r.stdout + r.stderr)
        sys.exit(f"git {' '.join(args)} failed")
    return r.stdout if capture else None


def main():
    ap = argparse.ArgumentParser(description="Publish the dashboard to GitHub Pages")
    ap.add_argument("--no-push", action="store_true", help="commit locally without pushing")
    ap.add_argument("-m", "--message", help="commit message")
    args = ap.parse_args()

    missing = [str(p) for p in FILES.values() if not p.exists()]
    if missing:
        sys.exit("Missing: " + ", ".join(missing))

    if not (REPO / ".git").exists():
        REPO.parent.mkdir(parents=True, exist_ok=True)
        r = subprocess.run(["git", "clone", REMOTE, str(REPO)])
        if r.returncode != 0:
            sys.exit("git clone failed")
    elif not args.no_push:
        git("pull", "--ff-only", "--quiet")

    for rel, src in FILES.items():
        dst = REPO / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dst)
    (REPO / ".nojekyll").touch()  # serve files as they are, without Jekyll
    (REPO / ".gitattributes").write_text("* -text\n")  # store files as they are: no line-ending warnings

    # Data files no longer built locally (e.g. a frequency that was dropped) leave the site
    tracked = git("ls-files", "data", capture=True).split()
    stale = [p for p in tracked if p not in FILES]
    if stale:
        git("rm", "--quiet", *stale)

    git("add", "--all")
    if not git("status", "--porcelain", capture=True).strip():
        print("Nothing changed since the last publish.")
        return
    stamp = datetime.now(SGT).strftime("%d %b %Y %H:%M SGT")
    git("commit", "--quiet", "-m", args.message or f"Refresh dashboard data, {stamp}")
    if args.no_push:
        print(f"Committed in {REPO}; not pushed.")
        return
    git("push", "--quiet", "origin", "HEAD")
    print(f"Pushed. GitHub Pages updates in a minute or two: {PAGES_URL}")


if __name__ == "__main__":
    main()
