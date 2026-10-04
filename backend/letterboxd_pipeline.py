"""
Letterboxd Data Pipeline
========================
Fetches a user's diary/ratings via their public RSS feed and exports
a clean dataset as CSV and Excel.

Usage:
    python letterboxd_pipeline.py <username>
    python letterboxd_pipeline.py dave --output my_movies
    python letterboxd_pipeline.py dave --no-dupes   # keep only latest rating per film

Dependencies:
    pip install requests beautifulsoup4 pandas openpyxl lxml
"""

import sys
import time
import argparse
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pathlib import Path
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_URL = "https://letterboxd.com"
HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}
DELAY = 0.5   # seconds between page requests


# ── Step 1: Fetch one RSS page ────────────────────────────────────────────────

def fetch_rss_page(username: str, page: int) -> BeautifulSoup | None:
    """
    Letterboxd RSS paginates via ?page=N.
    Returns a BeautifulSoup of the XML, or None on failure.
    """
    url = f"{BASE_URL}/{username}/rss/"
    params = {"page": page} if page > 1 else {}
    try:
        r = requests.get(url, headers=HEADERS, params=params, timeout=10)
        if r.status_code == 404:
            print(f"  ✗ User '{username}' not found.")
            return None
        if r.status_code != 200:
            print(f"  ✗ HTTP {r.status_code} on page {page}.")
            return None
        # Use lxml-xml parser for proper namespace handling
        return BeautifulSoup(r.text, "lxml-xml")
    except requests.RequestException as e:
        print(f"  ✗ Network error on page {page}: {e}")
        return None


# ── Step 2: Parse <item> entries ──────────────────────────────────────────────

def parse_items(soup: BeautifulSoup) -> list[dict]:
    """
    Each <item> in the RSS feed looks like:

        <item>
          <letterboxd:filmTitle>Toy Story 5</letterboxd:filmTitle>
          <letterboxd:filmYear>2026</letterboxd:filmYear>
          <letterboxd:memberRating>3.5</letterboxd:memberRating>
          <letterboxd:memberLike>Yes</letterboxd:memberLike>
          <letterboxd:watchedDate>2026-10-02</letterboxd:watchedDate>
          <letterboxd:rewatch>No</letterboxd:rewatch>
          <tmdb:movieId>1084244</tmdb:movieId>
          <link>https://letterboxd.com/dave/film/toy-story-5/</link>
          <pubDate>Sat, 3 Oct 2026 10:55:54 +1300</pubDate>
        </item>

    Items without a memberRating are diary entries with no star rating —
    we still include them but star_rating will be None.
    """
    records = []
    for item in soup.find_all("item"):

        def text(tag):
            node = item.find(tag)
            return node.get_text(strip=True) if node else None

        title      = text("filmTitle")
        year       = text("filmYear")
        watched    = text("watchedDate")
        rewatch    = text("rewatch")
        like       = text("memberLike")
        raw_rating = text("memberRating")
        tmdb_id    = text("movieId")
        link       = text("link")
        guid       = text("guid")

        # Derive slug from URL e.g. /dave/film/toy-story-5/ → toy-story-5
        slug = None
        if link:
            parts = [p for p in link.rstrip("/").split("/") if p]
            if "film" in parts:
                slug = parts[parts.index("film") + 1]

        records.append({
            "title":        title,
            "year":         int(year) if year else None,
            "watched_date": watched,
            "star_rating":  float(raw_rating) if raw_rating else None,
            "liked":        True if like == "Yes" else False,
            "rewatch":      True if rewatch == "Yes" else False,
            "tmdb_id":      tmdb_id,
            "slug":         slug,
            "url":          f"{BASE_URL}/film/{slug}/" if slug else link,
            "guid":         guid,
        })
    return records


# ── Step 3: Paginate until empty ──────────────────────────────────────────────

def run_pipeline(username: str, max_pages: int = 999, drop_dupes: bool = False) -> pd.DataFrame:
    """
    Walks every RSS page for the user and returns a combined DataFrame.

    The RSS feed is a diary feed (chronological watches), so rewatches
    appear as separate entries. Pass drop_dupes=True to keep only the
    most recent rating per film.
    """
    print(f"\n{'='*52}")
    print(f"  Letterboxd Pipeline  —  user: {username}")
    print(f"{'='*52}\n")

    all_records = []
    seen_guids = set()

    for page in range(1, max_pages + 1):
        print(f"  Fetching RSS page {page}…", end=" ", flush=True)
        soup = fetch_rss_page(username, page)
        if soup is None:
            break

        records = parse_items(soup)
        if not records:
            print("empty — done.")
            break

        # Filter to only records we haven't seen yet
        new_records = [r for r in records if r.get("guid") not in seen_guids]
        if not new_records:
            print("duplicate page — done.")
            break
        seen_guids.update(r["guid"] for r in new_records if r.get("guid"))
        all_records.extend(new_records)
        print(f"{len(records)} entries (running total: {len(all_records)})")
        time.sleep(DELAY)

    if not all_records:
        print("\n  No entries found. Is the profile public?\n")
        return pd.DataFrame()

    df = pd.DataFrame(all_records)
    df["watched_date"] = pd.to_datetime(df["watched_date"], errors="coerce")
    df = df.sort_values("watched_date", ascending=False).reset_index(drop=True)
    df = df.drop(columns=["guid"], errors="ignore")

    if drop_dupes:
        before = len(df)
        df = df.drop_duplicates(subset="slug", keep="first")
        print(f"\n  Dropped {before - len(df)} duplicate entries (rewatches).")

    # Summary
    rated = df["star_rating"].notna().sum()
    print(f"\n  ✓ {len(df)} total entries  |  {rated} with star ratings")
    if rated:
        print(f"    Avg rating : {df['star_rating'].mean():.2f} ★")
        print(f"    5★ films   : {(df['star_rating'] == 5.0).sum()}")
        print(f"    Liked      : {df['liked'].sum()}")

    return df


# ── Step 4: Export ────────────────────────────────────────────────────────────

def export_dataset(df: pd.DataFrame, username: str, stem: str | None = None):
    if df.empty:
        print("\n  Nothing to export.\n")
        return

    stem = stem or f"{username}_letterboxd_{datetime.now().strftime('%Y%m%d')}"

    # CSV
    csv_path = Path(stem + ".csv")
    df.to_csv(csv_path, index=False)
    print(f"\n  ✓ CSV  → {csv_path}")

    # Excel
    xlsx_path = Path(stem + ".xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ratings")
        ws = writer.sheets["Ratings"]

        header_fill = PatternFill("solid", fgColor="1F2C3D")
        header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        body_font   = Font(name="Arial", size=10)

        for col_idx, cell in enumerate(ws[1], 1):
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal="center")

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = body_font

        # Auto-width columns
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(ws.cell(row=r, column=col_idx).value or ""))
                for r in range(1, ws.max_row + 1)
            )
            ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

        ws.freeze_panes = "A2"

    print(f"  ✓ XLSX → {xlsx_path}\n")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Build a dataset of a Letterboxd user's watched films via RSS."
    )
    parser.add_argument("username", nargs="?")
    parser.add_argument("--username", "-u", dest="username_flag")
    parser.add_argument("--output",   "-o", help="Output filename stem (no extension)")
    parser.add_argument("--no-dupes", action="store_true",
                        help="Keep only the most recent entry per film (removes rewatches)")
    parser.add_argument("--max-pages", type=int, default=999,
                        help="Cap number of RSS pages (default: all)")
    args = parser.parse_args()

    username = args.username or args.username_flag
    if not username:
        parser.print_help()
        print("\n  Example: python letterboxd_pipeline.py dave\n")
        sys.exit(1)

    df = run_pipeline(
        username   = username.strip().lower(),
        max_pages  = args.max_pages,
        drop_dupes = args.no_dupes,
    )
    export_dataset(df, username, stem=args.output)


if __name__ == "__main__":
    main()