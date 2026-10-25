"""
YouTube Channel Transcript Extractor  v14.1  - PRODUCTION READY
=================================================================
Sirf channel URL do - baaki sab automatic.

OUTPUT (4 formats):
  Excel  (.xlsx)       - Daily use, filter/sort
  SQLite (.db)         - Search/query karo
  TSV    (.tsv)        - Backup / Google Sheets
  JSONL  (_llm.jsonl)  - GPT/Claude/Gemini analysis

v14.1 FIXES (vs v14.0):
  - FIXED: content_filter user choice was being overwritten
  - FIXED: OUTPUT_DIR/TEMP_DIR global shadowing in main()
  - FIXED: youtu.be URLs with hyphens now validated
  - FIXED: lang None handling in transcript source label
  - FIXED: langdetect import failure now logs warning
  - CLEANUP: removed unused imports & VideoInfo dataclass
"""
from __future__ import annotations

import subprocess
import sys
import shutil
import importlib.util
import os
import re
import json
import glob
import signal
import logging
import sqlite3
import threading
import asyncio
import time
import zipfile
import html
import hashlib
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from typing import Optional, List, Dict, Set, Tuple, Any
from pathlib import Path
from dataclasses import dataclass, field

# -- Dependency check ----------------------------------------------------
print("Checking system packages...")

_PKGS = {
    "yt_dlp": "yt-dlp",
    "youtube_transcript_api": "youtube-transcript-api",
    "openpyxl": "openpyxl",
    "nest_asyncio": "nest_asyncio",
    "langdetect": "langdetect",
    "tqdm": "tqdm",
}

_missing_pkgs = [pip for mod, pip in _PKGS.items() if importlib.util.find_spec(mod) is None]
if _missing_pkgs:
    print(f"   Installing missing packages: {', '.join(_missing_pkgs)}...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "--upgrade"] + _missing_pkgs, check=True)
    print("   Python packages ready!")
else:
    print("   Python packages already installed! (Skipped pip)")

if shutil.which("ffmpeg") is None:
    print("   ffmpeg missing - install it manually if yt-dlp subtitle download fails")

print("Setup Complete!\n")

# -- Google Drive mount -------------------------------------------------
_DRIVE_READY = False
_IN_COLAB = False

try:
    import google.colab  # noqa: F401
    _IN_COLAB = True
except ImportError:
    pass

if _IN_COLAB:
    print("Google Drive mount karna chahte ho?")
    print("   [1] Haan - files Drive mein save hongi  (recommended)")
    print("   [2] Nahi  - files ZIP mein download hongi")
    try:
        _drive_choice = input("   Choice (1/2, default=2): ").strip()
    except (EOFError, KeyboardInterrupt):
        _drive_choice = "2"

    if _drive_choice == "1":
        try:
            from google.colab import drive as _gdrive
            print("   Connecting Drive...")
            _gdrive.mount("/content/drive")
            _DRIVE_READY = True
            print("   Drive mounted!\n")
        except Exception as _e:
            print(f"   Drive mount fail: {_e}")
            print("   Fallback: ZIP download mode\n")

try:
    import nest_asyncio as _nest_asyncio
    _nest_asyncio.apply()
except ImportError:
    pass

try:
    from tqdm import tqdm
except ImportError:
    class tqdm:  # type: ignore
        def __init__(self, iterable=None, desc="", total=None, **kw):
            self.iterable = iterable
            self.desc = desc
            self.total = total or (len(iterable) if iterable else 0)
            self.n = 0
        def __iter__(self):
            for item in self.iterable:
                yield item
                self.n += 1
        def update(self, n=1): self.n += n
        def close(self): pass

# =====================================================================
#  CONFIG
# =====================================================================
MAX_VIDEOS       = 500
OUTPUT_DIR       = "/content/drive/MyDrive/YT_Transcripts" if _DRIVE_READY else "yt_output"
TEMP_DIR         = "/content/yt_temp" if _IN_COLAB else "yt_temp"
CHANNEL_URLS: list[str] = []
WORKERS          = 10
BATCH_SIZE       = 50
LLM_MAX_WORDS    = 500
RATE_LIMIT_DELAY = 2.0
MAX_RETRIES      = 3
RETRY_BACKOFF    = 2.0

_ANTIBLOCK_BASE: list[str] = [
    "--extractor-args", "youtube:player_client=tv_embedded",
    "--extractor-retries", "5",
    "--no-warnings",
]
_SCRAPE_SLEEP: list[str] = [
    "--sleep-requests", "1",
    "--sleep-interval", "2",
    "--max-sleep-interval", "4",
]
_YTDLP_LANG_PREF: list[str] = ["hi", "en", "en-US", "en-GB"]
EXCEL_MAX_ROWS = 1_048_576

STOP_FLAG: threading.Event = threading.Event()
ANTI_BLOCK_ARGS: list[str] = []

_channel_cache: Dict[str, str] = {}
_video_cache: Dict[str, list] = {}
_date_cache: Dict[str, str] = {}
_cache_lock = threading.Lock()

# =====================================================================
#  Data Classes
# =====================================================================
@dataclass
class ProcessingStats:
    total_videos: int = 0
    successful_transcripts: int = 0
    total_words: int = 0
    videos_count: int = 0
    shorts_count: int = 0
    errors: List[str] = field(default_factory=list)
    start_time: float = 0.0
    end_time: float = 0.0

logging.basicConfig(level=logging.WARNING, format="  [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# =====================================================================
#  yt-dlp helpers
# =====================================================================
def find_ytdlp() -> list[str]:
    path = shutil.which("yt-dlp") or shutil.which("yt-dlp.exe")
    if path:
        return [path]
    for folder in [
        os.path.join(os.path.dirname(sys.executable), "Scripts"),
        os.path.join(os.path.dirname(sys.executable)),
        os.path.join(sys.prefix, "Scripts"),
        os.path.join(sys.prefix, "bin"),
    ]:
        for name in ("yt-dlp.exe", "yt-dlp"):
            full = os.path.join(folder, name)
            if os.path.isfile(full):
                return [full]
    return [sys.executable, "-m", "yt_dlp"]

YTDLP: list[str] = find_ytdlp()

def run(args: list[str], timeout: int = 120,
        no_antiblock: bool = False, no_sleep: bool = False) -> tuple[str, str]:
    extra: list[str] = []
    if not no_antiblock:
        extra.extend(ANTI_BLOCK_ARGS)
        if not no_sleep:
            extra.extend(_SCRAPE_SLEEP)
    cmd = YTDLP + extra + args
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return r.stdout, r.stderr
    except subprocess.TimeoutExpired:
        log.warning("Timeout %ds: %s", timeout, " ".join(args[:4]))
        return "", ""
    except FileNotFoundError as e:
        log.error("yt-dlp nahi mila: %s", e)
        return "", str(e)

def setup_anti_block() -> None:
    global ANTI_BLOCK_ARGS
    ANTI_BLOCK_ARGS = list(_ANTIBLOCK_BASE)
    print("  Mode: tv_embedded + smart sleep")

# =====================================================================
#  Helpers
# =====================================================================
_WS_RE = re.compile(r"\s+")
_DATE_PATTERNS = [
    (re.compile(r"^(\d{4})-(\d{2})-(\d{2})"),
     lambda m: f"{m.group(3)}-{m.group(2)}-{m.group(1)}"),
    (re.compile(r"^(\d{8})$"),
     lambda m: f"{m.group(1)[6:8]}-{m.group(1)[4:6]}-{m.group(1)[0:4]}"),
]
_VTT_SKIP_RE = re.compile(r"^(WEBVTT|\d{2}:\d{2}|Kind:|Language:|NOTE|REGION|STYLE)")
_VTT_TAG_RE = re.compile(r"<[^>]+>")
_URL_PATTERN = re.compile(
    r"^https?://(?:www\.)?(?:"
    r"youtube\.com/(?:@[\w.-]+|channel/UC[\w-]{22}|c/[\w.-]+|user/[\w.-]+|shorts/[\w-]+)"
    r"|youtu\.be/[\w-]+"
    r")",
    re.IGNORECASE,
)

def clean(text: Optional[str]) -> str:
    return _WS_RE.sub(" ", text or "").strip()

def trim_for_llm(text: str, max_words: int = LLM_MAX_WORDS) -> str:
    if not text or text == "N/A":
        return text or "N/A"
    words = text.split()
    if len(words) <= max_words:
        return text
    half = max_words // 2
    return f"{' '.join(words[:half])} [...] {' '.join(words[-half:])}"

def validate_url(url: str) -> bool:
    if not url or not isinstance(url, str):
        return False
    url = url.strip()
    if "youtube.com" not in url and "youtu.be" not in url:
        return False
    return bool(_URL_PATTERN.match(url))

def safe_cache_get(cache: dict, key: str, lock: threading.Lock) -> Any:
    with lock:
        return cache.get(key)

def safe_cache_set(cache: dict, key: str, value: Any, lock: threading.Lock) -> None:
    with lock:
        cache[key] = value

# =====================================================================
#  URL & Options Input
# =====================================================================
def ask_urls() -> tuple[list[str], bool, str]:
    print("\n" + "=" * 60)
    print("  YouTube Channel Transcript Extractor  v14.1 - PRODUCTION READY")
    print("  Excel + SQLite + TSV + JSONL (LLM-ready)")
    print("=" * 60)
    print("  Channel URL daalo  (e.g. https://www.youtube.com/@channelname)")
    print("  Ek se zyada: ek-ek karke | Done: khali Enter")
    print("-" * 60)

    urls: list[str] = []
    while True:
        try:
            url = input(f"  URL {len(urls) + 1}: ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not url:
            if urls:
                break
            print("  Kam se kam ek URL daalo!")
            continue
        if not validate_url(url):
            print("  Valid YouTube URL chahiye!")
            continue
        urls.append(url)
        print(f"  Added ({len(urls)})")

    print("-" * 60)

    skip_date = True
    if urls:
        print("\n  Date Fetch:")
        print("     [1] Fetch karo  (slow)")
        print("     [2] SKIP karo   (fast) [DEFAULT]")
        try:
            c = input("  Choice (1/2, default=2): ").strip()
        except (EOFError, KeyboardInterrupt):
            c = "2"
        if c == "1":
            skip_date = False
            print("  Date fetch ON")
        else:
            print("  Date SKIPPED")

    print("-" * 60)

    content_filter = "both"
    if urls:
        print("\n  Content Filter:")
        print("     [1] Sirf Videos")
        print("     [2] Sirf Shorts")
        print("     [3] Both (Videos + Shorts) [DEFAULT]")
        try:
            f = input("  Choice (1/2/3, default=3): ").strip()
        except (EOFError, KeyboardInterrupt):
            f = "3"
        if f == "1":
            content_filter = "video"
            print("  Filter: Sirf Videos")
        elif f == "2":
            content_filter = "short"
            print("  Filter: Sirf Shorts")
        else:
            print("  Filter: Both")

    print("-" * 60)
    print("  Ctrl+C anytime - jo data mila woh save hoga")
    print("-" * 60)
    return urls, skip_date, content_filter

# =====================================================================
#  Channel ID
# =====================================================================
def get_channel_id(url: str) -> str:
    cached = safe_cache_get(_channel_cache, url, _cache_lock)
    if cached:
        return cached
    print("  Channel ID fetch kar raha hoon...")
    stdout, _ = run(["--flat-playlist", "--playlist-end", "1", "--print", "channel_id", url], no_sleep=True)
    cid = stdout.strip().split("\n")[0].strip() if stdout.strip() else ""
    if cid and cid.startswith("UC"):
        safe_cache_set(_channel_cache, url, cid, _cache_lock)
        return cid
    stdout2, _ = run(["--dump-single-json", "--flat-playlist", "--playlist-end", "1", url], no_sleep=True)
    try:
        data = json.loads(stdout2.strip().split("\n")[0])
        cid = data.get("channel_id") or data.get("uploader_id", "N/A")
        safe_cache_set(_channel_cache, url, cid, _cache_lock)
        return cid
    except Exception:
        pass
    safe_cache_set(_channel_cache, url, "N/A", _cache_lock)
    return "N/A"

def parse_date(val: Optional[str]) -> str:
    if not val or val in ("NA", "N/A", "None", "none", ""):
        return "N/A"
    val = val.strip()
    try:
        ts = int(float(val))
        if ts > 0:
            return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%d-%m-%Y")
    except (ValueError, OSError, OverflowError):
        pass
    for pattern, formatter in _DATE_PATTERNS:
        m = pattern.match(val)
        if m:
            try:
                return formatter(m)
            except (IndexError, ValueError):
                pass
    return "N/A"

def fetch_dates_batch(video_ids: List[str]) -> Dict[str, str]:
    if not video_ids or STOP_FLAG.is_set():
        return {}
    os.makedirs(TEMP_DIR, exist_ok=True)
    playlist_file = os.path.join(TEMP_DIR, "playlist.txt")
    urls = [f"https://www.youtube.com/watch?v={vid}" for vid in video_ids]
    with open(playlist_file, "w") as fh:
        fh.write("\n".join(urls))
    results: Dict[str, str] = {}
    try:
        stdout, _ = run([
            "--flat-playlist",
            "--print", "%(id)s\t%(upload_date)s\t%(timestamp)s\t%(release_date)s",
            "--playlist-items", f"1-{len(video_ids)}", "-a", playlist_file,
        ], timeout=300, no_sleep=True)
        for line in stdout.strip().split("\n"):
            line = line.strip()
            if not line or "\t" not in line:
                continue
            parts = line.split("\t")
            if len(parts) < 2:
                continue
            vid_id = parts[0].strip()
            date = "N/A"
            for i in range(1, len(parts)):
                d = parse_date(parts[i])
                if d != "N/A":
                    date = d
                    break
            results[vid_id] = date
            safe_cache_set(_date_cache, vid_id, date, _cache_lock)
    finally:
        try:
            os.remove(playlist_file)
        except OSError:
            pass
    return results

# =====================================================================
#  Video list
# =====================================================================
def fetch_tab(url: str, max_videos: int) -> tuple[list[dict], Set[str]]:
    args = [
        "--flat-playlist",
        "--print",
        "%(id)s\t%(title)s\t%(view_count)s\t%(timestamp)s\t%(upload_date)s\t%(webpage_url)s",
    ]
    if max_videos > 0:
        args += ["--playlist-end", str(max_videos)]
    args.append(url)
    stdout, _ = run(args, timeout=300)
    videos: list[dict] = []
    seen: Set[str] = set()
    for line in stdout.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        vid_id = parts[0].strip()
        if not vid_id or vid_id in seen:
            continue
        seen.add(vid_id)
        title = parts[1] if len(parts) > 1 else ""
        views = parts[2] if len(parts) > 2 else "0"
        ts_val = parts[3] if len(parts) > 3 else ""
        ud_val = parts[4] if len(parts) > 4 else ""
        wp_url = parts[5].strip() if len(parts) > 5 else ""
        vid_url = wp_url if wp_url.startswith("http") else f"https://www.youtube.com/watch?v={vid_id}"
        try:
            views_int = int(views)
        except ValueError:
            views_int = 0
        date = parse_date(ts_val)
        if date == "N/A":
            date = parse_date(ud_val)
        videos.append({
            "id": vid_id,
            "title": clean(title),
            "views": views_int,
            "views_M": round(views_int / 1_000_000, 2),
            "year": date,
            "url": vid_url,
            "type": "short" if "/shorts/" in vid_url else "video",
            "_date_pending": date == "N/A",
        })
    return videos, seen

def get_all_videos(channel_url: str, max_videos: int, content_filter: str) -> list[dict]:
    cache_key = f"{channel_url}_{content_filter}"
    cached = safe_cache_get(_video_cache, cache_key, _cache_lock)
    if cached:
        print("  Using cached video list...")
        return cached
    base = channel_url.rstrip("/")
    print(f"  Fetching content (Filter: {content_filter.upper()})...")
    all_videos: list[dict] = []
    all_ids: Set[str] = set()
    with ThreadPoolExecutor(max_workers=2) as exe:
        futures = {}
        if content_filter in ("video", "both"):
            futures["video"] = exe.submit(fetch_tab, base + "/videos", max_videos)
        if content_filter in ("short", "both"):
            futures["short"] = exe.submit(fetch_tab, base + "/shorts", max_videos)
        if "video" in futures:
            vids, ids1 = futures["video"].result()
            all_videos.extend(vids)
            all_ids.update(ids1)
            print(f"     -> {len(vids)} videos found")
        if "short" in futures:
            shorts, ids2 = futures["short"].result()
            new_shorts = [s for s in shorts if s["id"] not in all_ids]
            all_videos.extend(new_shorts)
            all_ids.update(ids2)
            print(f"     -> {len(new_shorts)} shorts found")
    safe_cache_set(_video_cache, cache_key, all_videos, _cache_lock)
    return all_videos

# =====================================================================
#  Transcript helpers
# =====================================================================
def detect_lang_from_title(title: str) -> str:
    if not title or len(title.strip()) < 4:
        return "en"
    try:
        from langdetect import detect
        return detect(title.strip())
    except Exception as e:
        log.debug("langdetect fail: %s", e)
        return "en"

def _score_transcript(t, detected_lang: str) -> int:
    score = 0
    lang_code = getattr(t, "language_code", "") or ""
    is_generated = getattr(t, "is_generated", True)
    is_translatable = getattr(t, "is_translatable", False)
    if not is_generated:
        score += 20
    if lang_code == detected_lang:
        score += 10
    elif lang_code.startswith(detected_lang[:2]):
        score += 6
    if is_generated:
        score += 5
    if is_translatable and lang_code[:2] != detected_lang[:2]:
        score -= 5
    return score

def transcript_yta_auto(video_id: str, title: str = "") -> tuple[Optional[str], str]:
    if STOP_FLAG.is_set():
        return None, ""
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
    except ImportError:
        return None, ""
    detected = detect_lang_from_title(title)
    try:
        tlist = YouTubeTranscriptApi.list_transcripts(video_id)
        available = list(tlist)
        if not available:
            return None, ""
        scored = sorted([(_score_transcript(t, detected), t) for t in available], key=lambda x: x[0], reverse=True)
        for sc, tr in scored:
            if sc <= 0:
                break
            try:
                data = tr.fetch()
                text = " ".join(
                    (s.text if hasattr(s, "text") else s.get("text", "")).strip() for s in data
                ).strip()
                if text:
                    lang = getattr(tr, "language_code", None) or detected
                    return text, lang
            except Exception:
                continue
    except Exception as e:
        log.warning("yta_auto fail %s: %s", video_id, e)
    return None, ""

def _parse_vtt(filepath: str) -> str:
    try:
        with open(filepath, encoding="utf-8") as fh:
            content = fh.read()
    except OSError:
        return ""
    result: list[str] = []
    for line in content.split("\n"):
        line = line.strip()
        if not line or _VTT_SKIP_RE.match(line) or "-->" in line:
            continue
        cleaned = html.unescape(_VTT_TAG_RE.sub("", line)).strip()
        if cleaned and (not result or result[-1] != cleaned):
            result.append(cleaned)
    return " ".join(result)

def transcript_ytdlp(video_id: str) -> Optional[str]:
    if STOP_FLAG.is_set():
        return None
    os.makedirs(TEMP_DIR, exist_ok=True)
    out_tpl = os.path.join(TEMP_DIR, video_id)
    try:
        run([
            "--write-auto-sub", "--write-sub",
            "--sub-langs", "hi,en,en-US,en-GB,en.*",
            "--skip-download", "--sub-format", "vtt",
            "--output", out_tpl,
            f"https://www.youtube.com/watch?v={video_id}",
        ], timeout=60, no_sleep=True)
        files = glob.glob(os.path.join(TEMP_DIR, f"{video_id}*.vtt"))
        if not files:
            return None
        chosen = files[0]
        for lang in _YTDLP_LANG_PREF:
            for f in files:
                if f".{lang}." in f or f"-{lang}." in f:
                    chosen = f
                    break
            else:
                continue
            break
        return _parse_vtt(chosen) or None
    finally:
        for f in glob.glob(os.path.join(TEMP_DIR, f"{video_id}*")):
            try:
                os.remove(f)
            except OSError:
                pass

def get_transcript_with_retry(video_id: str, title: str = "", is_short: bool = False) -> tuple[str, str]:
    if STOP_FLAG.is_set():
        return "N/A", "stopped"
    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            text, lang = transcript_yta_auto(video_id, title)
            if text:
                return text, f"yt-api:{lang or 'auto'}"
            text = transcript_ytdlp(video_id)
            if text and text.strip():
                return text.strip(), "yt-dlp"
            return "N/A", "none"
        except Exception as e:
            last_error = str(e)
            log.warning("Attempt %d/%d failed for %s: %s", attempt + 1, MAX_RETRIES, video_id, e)
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_BACKOFF ** attempt)
    log.error("All %d attempts failed for %s: %s", MAX_RETRIES, video_id, last_error)
    return "N/A", f"error:{last_error}"

async def fetch_transcripts_async(videos: List[dict], workers: int = WORKERS) -> List[Tuple[str, str]]:
    semaphore = asyncio.Semaphore(workers)
    results: List[Tuple[str, str]] = [("N/A", "none")] * len(videos)

    async def fetch_one(idx: int, video: dict) -> None:
        async with semaphore:
            if STOP_FLAG.is_set():
                return
            loop = asyncio.get_running_loop()
            tr, src = await loop.run_in_executor(
                None, get_transcript_with_retry, video["id"], video.get("title", ""),
                video.get("type") == "short",
            )
            tag = "SHORT" if video.get("type") == "short" else "VIDEO"
            ok = "OK" if tr != "N/A" else "FAIL"
            print(f"  [{idx+1:>3}/{len(videos)}] {tag:<5} {video['title'][:43]:<43} views={video['views_M']}M  {ok} {src}")
            results[idx] = (tr, src)

    await asyncio.gather(*(fetch_one(i, v) for i, v in enumerate(videos)), return_exceptions=True)
    return results

# =====================================================================
#  Output formats
# =====================================================================
_CELL_LIMIT = 32_000

def _split_chunks(text: str) -> list[str]:
    if not text or text == "N/A":
        return [text or "N/A"]
    chunks: list[str] = []
    while len(text) > _CELL_LIMIT:
        cut = text.rfind(" ", 0, _CELL_LIMIT)
        if cut == -1:
            cut = _CELL_LIMIT
        chunks.append(text[:cut].strip())
        text = text[cut:].strip()
    if text:
        chunks.append(text)
    return chunks

def make_excel(rows: list[dict], filepath: str) -> bool:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return False
    try:
        if len(rows) > EXCEL_MAX_ROWS:
            log.error("Too many rows (%d) for Excel. Max is %d.", len(rows), EXCEL_MAX_ROWS)
            rows = rows[:EXCEL_MAX_ROWS]
            print(f"  Truncated to {EXCEL_MAX_ROWS} rows (Excel limit)")
        wb = Workbook()
        ws = wb.active
        ws.title = "YouTube Data"
        all_chunks = [_split_chunks(r["transcript"]) for r in rows]
        max_chunks = max((len(c) for c in all_chunks), default=1)
        base_hdrs = ["#", "Channel", "Title", "Type", "Views (M)", "Upload Date", "URL"]
        tr_hdrs = [f"Transcript_{i+1}" if max_chunks > 1 else "Transcript" for i in range(max_chunks)]
        headers = base_hdrs + tr_hdrs + ["Source"]
        hfill = PatternFill("solid", start_color="1F4E79")
        hfont = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        afill = PatternFill("solid", start_color="EBF3FB")
        bside = Side(style="thin", color="BFBFBF")
        brd = Border(left=bside, right=bside, top=bside, bottom=bside)
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = hfont; c.fill = hfill; c.border = brd
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28
        normal_font = Font(name="Arial", size=10)
        link_font = Font(name="Arial", size=10, color="0563C1", underline="single")
        normal_align = Alignment(vertical="top")
        wrap_align = Alignment(vertical="top", wrap_text=True)
        for i, (row, chunks) in enumerate(zip(rows, all_chunks), 2):
            while len(chunks) < max_chunks:
                chunks.append("")
            base_vals = [row["sr"], row["channel"], row["title"], row.get("type", "video"),
                         row["views_M"], row["year"], row["url"]]
            for col, val in enumerate(base_vals, 1):
                c = ws.cell(row=i, column=col, value=val)
                c.font = normal_font; c.border = brd
                c.alignment = wrap_align if col in (3, 7) else normal_align
                if i % 2 == 0:
                    c.fill = afill
                if col == 7 and val not in ("N/A", ""):
                    c.hyperlink = val
                    c.font = link_font
            for ci, chunk_text in enumerate(chunks):
                col = len(base_hdrs) + ci + 1
                c = ws.cell(row=i, column=col, value=chunk_text)
                c.font = normal_font; c.border = brd; c.alignment = wrap_align
                if i % 2 == 0:
                    c.fill = afill
            src_col = len(base_hdrs) + max_chunks + 1
            c = ws.cell(row=i, column=src_col, value=row["source"])
            c.font = normal_font; c.border = brd; c.alignment = normal_align
            if i % 2 == 0:
                c.fill = afill
        widths = [5, 18, 35, 8, 12, 12, 40] + [60] * max_chunks + [14]
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
        wb.save(filepath)
        return True
    except Exception as e:
        log.error("Excel error: %s", e)
        return False

def make_sqlite(rows: list[dict], filepath: str) -> bool:
    try:
        conn = sqlite3.connect(filepath)
        cur = conn.cursor()
        cur.execute("""CREATE TABLE IF NOT EXISTS videos (
            sr INTEGER PRIMARY KEY, channel TEXT, channel_id TEXT, title TEXT, type TEXT,
            views_M REAL, views INTEGER, upload_date TEXT, url TEXT, transcript TEXT,
            source TEXT, word_count INTEGER)""")
        cur.execute("CREATE VIRTUAL TABLE IF NOT EXISTS videos_fts USING fts5"
                    "(title, transcript, content='videos', content_rowid='sr')")
        cur.execute("DELETE FROM videos")
        cur.execute("DELETE FROM videos_fts")
        data = [
            (r["sr"], r["channel"], r["channel_id"], r["title"], r.get("type", "video"),
             r["views_M"], r["views"], r["year"], r["url"], r["transcript"], r["source"],
             len(r["transcript"].split()) if r["transcript"] != "N/A" else 0)
            for r in rows
        ]
        cur.executemany("INSERT INTO videos VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", data)
        fts_data = [(r["sr"], r["title"], r["transcript"]) for r in rows if r["transcript"] != "N/A"]
        if fts_data:
            cur.executemany("INSERT INTO videos_fts (rowid, title, transcript) VALUES (?,?,?)", fts_data)
            cur.execute("INSERT INTO videos_fts(videos_fts) VALUES('rebuild')")
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        log.error("SQLite error: %s", e)
        return False

def make_tsv(rows: list[dict], filepath: str) -> bool:
    try:
        hdrs = ["Sr", "Channel", "Title", "Type", "Views_M", "Upload_Date", "URL", "Transcript", "Source"]
        with open(filepath, "w", encoding="utf-8-sig") as fh:
            fh.write("\t".join(hdrs) + "\n")
            for r in rows:
                tr = (r["transcript"].replace("\n", " ").replace("\t", "  ")
                      if r["transcript"] != "N/A" else "N/A")
                fh.write("\t".join([
                    str(r["sr"]), r["channel"], r["title"], r.get("type", "video"),
                    str(r["views_M"]), r["year"], r["url"], tr, r["source"],
                ]) + "\n")
        return True
    except Exception as e:
        log.error("TSV error: %s", e)
        return False

def make_jsonl_llm(rows: list[dict], filepath: str) -> bool:
    try:
        with open(filepath, "w", encoding="utf-8") as fh:
            for r in rows:
                fh.write(json.dumps({
                    "t": r["title"], "v": r["views_M"], "d": r["year"],
                    "tp": r.get("type", "video"),
                    "tr": trim_for_llm(r["transcript"], LLM_MAX_WORDS),
                }, ensure_ascii=False) + "\n")
        return True
    except Exception as e:
        log.error("JSONL error: %s", e)
        return False

# =====================================================================
#  Misc helpers
# =====================================================================
def url_to_filename(url: str) -> str:
    u = url.rstrip("/")
    m = re.search(r"/@([\w.-]+)", u)
    if m:
        return "@" + m.group(1)
    m = re.search(r"/channel/(UC[\w-]+)", u)
    if m:
        return m.group(1)
    m = re.search(r"/(?:c|user)/([\w.-]+)", u)
    if m:
        return "c_" + m.group(1)
    h = hashlib.md5(u.encode()).hexdigest()[:8]
    return f"yt_{h}"

def get_combined_filename(urls: list[str]) -> str:
    if len(urls) == 1:
        return url_to_filename(urls[0])
    return f"{url_to_filename(urls[0])}_plus_{len(urls) - 1}_more"

def _handle_sigint(sig, frame) -> None:
    if not STOP_FLAG.is_set():
        STOP_FLAG.set()
        print("\n\n  Ctrl+C! Tasks khatam ho rahe hain, phir save...\n")
    signal.signal(signal.SIGINT, signal.SIG_DFL)

def download_as_zip(file_paths: list[str], zip_name: str) -> None:
    existing = [p for p in file_paths if os.path.isfile(p)]
    if not existing:
        return
    zip_path = os.path.join(os.path.dirname(existing[0]), zip_name)
    total = len(existing)
    print(f"\n  ZIP bana raha hoon: {zip_name}")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for i, fp in enumerate(existing, 1):
            arcname = os.path.basename(fp)
            zf.write(fp, arcname)
            pct = i * 100 // total
            bar = "#" * (pct * 30 // 100) + "-" * (30 - pct * 30 // 100)
            print(f"\r     [{bar}] {i}/{total}  {arcname} "
                  f"({os.path.getsize(fp) / 1024:.0f} KB)", end="", flush=True)
    print(f"\n  ZIP ready: {zip_path}  ({os.path.getsize(zip_path) / 1024:.0f} KB)")
    if _IN_COLAB:
        try:
            from google.colab import files as _colab_files
            print("  Browser download shuru ho raha hai...")
            _colab_files.download(zip_path)
        except Exception as e:
            print(f"  Auto-download fail: {e}")
    for fp in existing:
        try:
            os.remove(fp)
        except OSError:
            pass

# =====================================================================
#  MAIN
# =====================================================================
def main() -> None:
    global OUTPUT_DIR, TEMP_DIR

    signal.signal(signal.SIGINT, _handle_sigint)
    print(f"\n  yt-dlp: {' '.join(YTDLP)}")
    ver, _ = run(["--version"], timeout=10, no_antiblock=True)
    if not ver.strip():
        sys.exit("yt-dlp nahi chala!")
    print(f"  yt-dlp: {ver.strip()}")

    setup_anti_block()

    if CHANNEL_URLS:
        urls, skip_date, content_filter = CHANNEL_URLS, True, "both"
    else:
        urls, skip_date, content_filter = ask_urls()

    if not urls:
        sys.exit("Koi URL nahi diya.")

    if not _DRIVE_READY:
        OUTPUT_DIR = str(Path(OUTPUT_DIR).resolve())
    TEMP_DIR = str(Path(TEMP_DIR).resolve())

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)

    all_rows: list[dict] = []
    sr = 1
    stats = ProcessingStats(start_time=time.time())

    try:
        for idx, ch_url in enumerate(tqdm(urls, desc="Processing channels")):
            if STOP_FLAG.is_set():
                break
            if idx > 0:
                print(f"  Rate limiting: {RATE_LIMIT_DELAY}s delay...")
                time.sleep(RATE_LIMIT_DELAY)
            ch_url = ch_url.strip().rstrip("/")
            print(f"\n{'=' * 60}\n{ch_url}\n{'=' * 60}")
            channel_id = get_channel_id(ch_url)
            print(f"  Channel ID: {channel_id}")
            videos = get_all_videos(ch_url, MAX_VIDEOS, content_filter)
            if not videos:
                print("  Koi video/short nahi mila.")
                continue
            stats.total_videos += len(videos)
            if not skip_date:
                na_vids = [v for v in videos if v.get("_date_pending")]
                if na_vids:
                    print(f"  Fetching dates for {len(na_vids)} items...", flush=True)
                    for bs in range(0, len(na_vids), BATCH_SIZE):
                        if STOP_FLAG.is_set():
                            break
                        batch = na_vids[bs:bs + BATCH_SIZE]
                        dates = fetch_dates_batch([v["id"] for v in batch])
                        for v in batch:
                            v["year"] = dates.get(v["id"], "N/A")
                            v["_date_pending"] = False
                        pct = min(bs + BATCH_SIZE, len(na_vids)) * 100 // len(na_vids)
                        bar = "#" * (pct * 30 // 100) + "-" * (30 - pct * 30 // 100)
                        print(f"\r     [{bar}] {min(bs+BATCH_SIZE, len(na_vids))}/{len(na_vids)} "
                              f"({pct}%)", end="", flush=True)
                    print("\n  Date fetch complete" if not STOP_FLAG.is_set() else "\n  Stopped.")

            videos.sort(key=lambda x: x["views"], reverse=True)
            ch_name = re.sub(
                r"https?://(www\.)?youtube\.com/(@|channel/|c/|user/)?", "", ch_url
            ).strip("/")

            print(f"  {len(videos)} items - Async + Auto lang detect")
            results = asyncio.run(fetch_transcripts_async(videos, WORKERS))

            for i, v in enumerate(videos):
                tr, src = results[i] if results[i] else ("N/A", "none")
                v.pop("_date_pending", None)
                if tr != "N/A":
                    stats.successful_transcripts += 1
                    stats.total_words += len(tr.split())
                if v.get("type") == "short":
                    stats.shorts_count += 1
                else:
                    stats.videos_count += 1
                if src.startswith("error:"):
                    stats.errors.append(f"{v['id']}: {src}")
                all_rows.append({
                    "sr": sr, "channel": ch_name, "channel_id": channel_id,
                    "title": v["title"], "views": v["views"], "views_M": v["views_M"],
                    "year": v["year"], "url": v["url"], "type": v.get("type", "video"),
                    "transcript": tr, "source": src,
                })
                sr += 1
    except KeyboardInterrupt:
        STOP_FLAG.set()
        print("\n\n  Ctrl+C - saving...\n")

    if not all_rows:
        return print("\nKoi data nahi mila.")

    all_rows.sort(key=lambda x: x["views"], reverse=True)
    for i, r in enumerate(all_rows, 1):
        r["sr"] = i

    base = re.sub(r'[<>:"/\\|?*]', "_", get_combined_filename(urls))
    if STOP_FLAG.is_set():
        base += "_partial"

    saved_files: list[str] = []
    print("\n  Saving outputs...\n")
    for fmt, ext, func in [
        ("Excel", ".xlsx", make_excel),
        ("SQLite", ".db", make_sqlite),
        ("TSV", ".tsv", make_tsv),
        ("JSONL", "_llm.jsonl", make_jsonl_llm),
    ]:
        fp = os.path.join(OUTPUT_DIR, f"{base}{ext}")
        if func(all_rows, fp):
            saved_files.append(fp)
            print(f"  {fmt:<7}: {os.path.basename(fp)} ({os.path.getsize(fp)/1024:.0f} KB)")
        else:
            print(f"  {fmt} save fail hua.")

    if not _DRIVE_READY and saved_files:
        download_as_zip(saved_files, f"{base}_transcripts.zip")

    stats.end_time = time.time()
    elapsed = stats.end_time - stats.start_time

    print(f"\n{'=' * 60}\n  {'PARTIAL' if STOP_FLAG.is_set() else 'COMPLETE'}")
    print(f"  Filter Used: {content_filter.upper()}")
    print(f"  Total      : {len(all_rows)}  "
          f"(videos={stats.videos_count}  shorts={stats.shorts_count})")
    rate = (stats.successful_transcripts * 100 // len(all_rows)) if all_rows else 0
    print(f"  Transcript : {stats.successful_transcripts}/{len(all_rows)}  ({rate}%)")
    print(f"  Words      : {stats.total_words:,}")
    print(f"  Time       : {elapsed:.1f}s")
    if stats.errors:
        print(f"  Errors     : {len(stats.errors)} (check logs)")
    print(f"{'=' * 60}\n")

    for f in glob.glob(os.path.join(TEMP_DIR, "*")):
        try:
            os.remove(f)
        except OSError:
            pass

if __name__ == "__main__":
    main()