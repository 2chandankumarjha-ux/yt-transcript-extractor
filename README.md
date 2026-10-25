# YouTube Channel Transcript Extractor

[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Extract transcripts from an entire YouTube channel in one command, and export to four formats ready for analysis, search, or AI pipelines.

**v14.1** - Production-ready single-file script. Works in Google Colab and locally.

## Features

- Bulk extraction - transcripts from hundreds of videos at once
- Auto language detection - Hindi, English, and more (via `langdetect`)
- Shorts + long videos - filter or grab both
- Two transcript sources - YouTube Transcript API first, `yt-dlp` fallback
- Async + thread-safe - configurable workers, cached, rate-limited
- Anti-block - `tv_embedded` client + smart sleep
- 4 output formats: Excel, SQLite (FTS5), TSV, JSONL (LLM-ready)
- Google Drive - optional save to `MyDrive/YT_Transcripts/`
- ZIP auto-download - when Drive is not connected (Colab)
- Ctrl+C safe - partial results always saved
- LLM-optimized - `trim_for_llm()` keeps head+tail per transcript

## Quick Start (Colab)

1. Open Google Colab
2. Paste `yt_transcript_extractor.py` into a cell or upload it
3. Run `!python yt_transcript_extractor.py`
4. Follow the prompts

## Local Install

Prerequisites: Python 3.9+ and ffmpeg.

```bash
git clone https://github.com/USERNAME/yt-transcript-extractor.git
cd yt-transcript-extractor
pip install -r requirements.txt
python yt_transcript_extractor.py
```

## Output Formats

- **Excel** - filter, sort, hyperlinks, frozen header. Transcripts over 32,000 chars split across columns.
- **SQLite** - FTS5 full-text search:
  ```sql
  SELECT title FROM videos_fts WHERE videos_fts MATCH 'quantum';
  ```
- **TSV** - UTF-8 with BOM, opens cleanly in Sheets/Excel.
- **JSONL** - one object per line:
  ```json
  {"t":"title","v":1.2,"d":"15-01-2024","tp":"video","tr":"head [...] tail"}
  ```

## Configuration

| Constant | Default | Description |
|----------|---------|-------------|
| `MAX_VIDEOS` | 500 | Max videos per channel |
| `WORKERS` | 10 | Concurrent fetches |
| `BATCH_SIZE` | 50 | Date-fetch batch size |
| `LLM_MAX_WORDS` | 500 | JSONL transcript limit |
| `RATE_LIMIT_DELAY` | 2.0 | Seconds between channels |
| `MAX_RETRIES` | 3 | Retries per transcript |
| `CHANNEL_URLS` | [] | Non-interactive mode |

## Troubleshooting

| Problem | Fix |
|---------|-----|
| yt-dlp not running | `pip install -U yt-dlp` |
| All transcripts N/A | Captions disabled or IP blocked - try VPN |
| ffmpeg not found | Install ffmpeg |
| Rate limited | Increase `RATE_LIMIT_DELAY` |
| asyncio error in Jupyter | `pip install nest-asyncio` |

## Legal

Personal research, accessibility, and archival use only. Respect YouTube ToS, creator copyright, and rate limits.

## License

MIT - see [LICENSE](LICENSE).