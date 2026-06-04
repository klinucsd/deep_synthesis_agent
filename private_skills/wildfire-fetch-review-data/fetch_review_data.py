"""
fetch_review_data.py
====================

Helper for the wildfire-fetch-review-data skill. Downloads the wildfire
review-data bundle from the Fire_Risk_Review Google Drive folder using the
reviewer's OAuth token.

Three artifacts land in the output directory (current working dir by default):

  - missing_rubrics/                                           (unzipped from missing_rubrics.zip)
  - Preliminary-Extract-Science-Technical-Evaluator-Feedback-20260603.xlsx
  - Preliminary-Extract-User-Practitioner-Evaluation-Feedback-20260603.xlsx

Drive IDs are baked into this private skill (project-specific data, not for the
public notebook). If the source folder is reorganized, update the constants
below.

Public entry point: `fetch_review_data(output_dir=".", force=False)`.
"""

from __future__ import annotations

import io
import json
import os
import sys
import zipfile
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Drive file IDs — private to this skill, baked in for the Fire_Risk_Review folder.
# Folder URL: https://drive.google.com/drive/folders/1nVohBOQDlh3qm32XBYbY25oVkFBeYR5-
# ---------------------------------------------------------------------------

_FOLDER_ID = "1nVohBOQDlh3qm32XBYbY25oVkFBeYR5-"

_FILES = [
    {
        "drive_id": "1uj4gU2oBW2nkN0NDCczdP0T9c9FmwEnu",
        "local_name": "missing_rubrics.zip",
        "kind": "zip",
        "extract_to": "missing_rubrics",
        "expected_subdirs": 12,
    },
    {
        "drive_id": "1800nf8CdknMZlCyHQaIIa22lGGg9GUts",
        "local_name": "Preliminary-Extract-Science-Technical-Evaluator-Feedback-20260603.xlsx",
        "kind": "file",
    },
    {
        "drive_id": "1TeojGvWMDMi6V_nBMviXNBYNIWHsrf84",
        "local_name": "Preliminary-Extract-User-Practitioner-Evaluation-Feedback-20260603.xlsx",
        "kind": "file",
    },
]


# ---------------------------------------------------------------------------
# Token discovery
# ---------------------------------------------------------------------------

_TOKEN_SEARCH_PATHS = [
    os.environ.get("GDRIVE_TOKEN", ""),
    "/home/jovyan/work/_User-Persistent-Storage_CephBlock_/.gdrive_token.json",
    str(Path.home() / ".gdrive_token.json"),
    "./.gdrive_token.json",
]


def _find_token() -> Path:
    for candidate in _TOKEN_SEARCH_PATHS:
        if candidate and Path(candidate).is_file():
            return Path(candidate)
    raise FileNotFoundError(
        "Could not find .gdrive_token.json. Searched: "
        + ", ".join(p for p in _TOKEN_SEARCH_PATHS if p)
    )


def _build_drive_service():
    """Construct an authenticated Drive v3 client from the saved token."""
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    token_path = _find_token()
    t = json.loads(token_path.read_text())
    creds = Credentials(
        token=t.get("token"),
        refresh_token=t.get("refresh_token"),
        token_uri=t.get("token_uri"),
        client_id=t.get("client_id"),
        client_secret=t.get("client_secret"),
        scopes=t.get("scopes"),
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False), token_path


# ---------------------------------------------------------------------------
# Download primitives
# ---------------------------------------------------------------------------

def _file_size(svc, drive_id: str) -> int:
    meta = svc.files().get(fileId=drive_id, fields="size,name").execute()
    return int(meta.get("size", 0))


def _download_one(svc, drive_id: str, dest: Path, progress_label: str = ""):
    """Stream a Drive file to dest with a simple progress indicator."""
    from googleapiclient.http import MediaIoBaseDownload

    dest.parent.mkdir(parents=True, exist_ok=True)
    request = svc.files().get_media(fileId=drive_id)
    with dest.open("wb") as fh:
        downloader = MediaIoBaseDownload(fh, request, chunksize=8 * 1024 * 1024)
        done = False
        last_print = -1
        while not done:
            status, done = downloader.next_chunk()
            if status is not None:
                pct = int(status.progress() * 100)
                if pct >= last_print + 5:           # log every ~5%
                    print(f"  {progress_label}  {pct}%")
                    sys.stdout.flush()
                    last_print = pct
    print(f"  {progress_label}  done ({dest.stat().st_size:,} bytes)")
    sys.stdout.flush()


# ---------------------------------------------------------------------------
# Zip extraction
# ---------------------------------------------------------------------------

def _zip_already_extracted(extract_root: Path, expected_subdirs: int) -> bool:
    if not extract_root.is_dir():
        return False
    actual = sum(1 for p in extract_root.iterdir() if p.is_dir() and p.name.startswith("Deep_Synthesis__"))
    return actual >= expected_subdirs


def _extract_zip(zip_path: Path, output_dir: Path, expected_subdirs: int):
    """Extract a zip whose internal layout is `missing_rubrics/Deep_Synthesis__*/...`.
    Result: `<output_dir>/missing_rubrics/Deep_Synthesis__*/...` ready for downstream skills."""
    print(f"  extracting {zip_path.name} → {output_dir}/")
    sys.stdout.flush()
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(output_dir)
    final = output_dir / "missing_rubrics"
    if not final.is_dir():
        # Some zip producers omit the top-level directory; reconstruct it.
        raise RuntimeError(
            f"Zip {zip_path} did not contain the expected top-level "
            f"`missing_rubrics/` directory."
        )
    n = sum(1 for p in final.iterdir() if p.is_dir() and p.name.startswith("Deep_Synthesis__"))
    if n < expected_subdirs:
        print(f"  WARNING: expected {expected_subdirs} Deep_Synthesis__ subfolders, found {n}")
        sys.stdout.flush()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def fetch_review_data(output_dir: str | Path = ".", force: bool = False) -> dict:
    """Download the wildfire-review data bundle.

    Parameters
    ----------
    output_dir : str | Path
        Directory the artifacts land in. Defaults to the current working directory.
    force : bool
        If True, re-download even when local copies match the Drive sizes.

    Returns
    -------
    dict
        {"missing_rubrics_dir": <Path>,
         "science_xlsx": <Path>,
         "practitioner_xlsx": <Path>,
         "token_used": <Path>}
    """
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"output directory: {output_dir}")
    sys.stdout.flush()

    svc, token_path = _build_drive_service()
    print(f"token: {token_path}")
    sys.stdout.flush()

    science_xlsx_path: Optional[Path] = None
    practitioner_xlsx_path: Optional[Path] = None
    missing_rubrics_dir: Optional[Path] = None

    for spec in _FILES:
        local = output_dir / spec["local_name"]
        label = spec["local_name"]

        # Decide whether to skip.
        skip_download = False
        if not force and local.is_file():
            try:
                remote_size = _file_size(svc, spec["drive_id"])
                if local.stat().st_size == remote_size:
                    print(f"skip (already present, size matches): {label}")
                    sys.stdout.flush()
                    skip_download = True
            except Exception:
                pass  # fall through to re-download on any metadata error

        if not skip_download:
            print(f"downloading: {label}")
            sys.stdout.flush()
            _download_one(svc, spec["drive_id"], local, progress_label=label)

        # Post-processing
        if spec["kind"] == "zip":
            extract_root = output_dir / spec["extract_to"]
            if force or not _zip_already_extracted(extract_root, spec["expected_subdirs"]):
                _extract_zip(local, output_dir, spec["expected_subdirs"])
            else:
                print(f"skip extraction (already present): {extract_root}")
                sys.stdout.flush()
            missing_rubrics_dir = extract_root
        else:
            if "Science" in spec["local_name"]:
                science_xlsx_path = local
            elif "Practitioner" in spec["local_name"] or "User" in spec["local_name"]:
                practitioner_xlsx_path = local

    return {
        "missing_rubrics_dir": missing_rubrics_dir,
        "science_xlsx": science_xlsx_path,
        "practitioner_xlsx": practitioner_xlsx_path,
        "token_used": token_path,
    }


# ---------------------------------------------------------------------------
# CLI

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Fetch wildfire review data from Drive.")
    ap.add_argument("--output-dir", default=".")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()
    out = fetch_review_data(output_dir=args.output_dir, force=args.force)
    for k, v in out.items():
        print(f"{k}: {v}")
