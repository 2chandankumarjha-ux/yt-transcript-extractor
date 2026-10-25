# Contributing

Thanks for your interest!

## Reporting Bugs

Open an issue with:
1. What happened vs what you expected
2. Full traceback (if any)
3. Python version (`python --version`)
4. `yt-dlp --version`
5. OS

## Pull Requests

1. Fork the repo and create a branch: `git checkout -b fix/my-bug`
2. Make changes; test on at least one public channel
3. Run linters:
   ```bash
   pip install ruff black
   ruff check yt_transcript_extractor.py
   black --check yt_transcript_extractor.py
   ```
4. Commit: `git commit -m "fix: content_filter now respects user choice"`
5. Push and open a PR with a clear description

## Code Style

- PEP 8, line length 100
- Type hints for public functions
- Keep it a single file - zero-setup deployment is the goal

## Commit Messages

- `feat:` new feature
- `fix:` bug fix
- `docs:` documentation
- `refactor:` code change that neither fixes a bug nor adds a feature
- `chore:` tooling / dependencies