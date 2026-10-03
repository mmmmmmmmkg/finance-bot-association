# Finance Telegram Bot

A Telegram bot for recording payment requests with receipts, reviewing payments, managing projects and expenses, viewing financial reports, and exporting reports to Excel. The bot uses SQLite and Telegram long polling.

## Local setup

Requirements: Python 3.12 and the packages pinned in `requirements.txt`.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
# Fill the local environment values in .env, then run:
python bot.py
```

## Environment variables

- `BOT_TOKEN` — required Telegram bot token.
- `ADMIN_IDS` — required comma-separated Telegram administrator IDs.
- `DB_PATH` — optional SQLite file path; defaults to `finance_bot.db` in the working directory.

Keep real values in environment settings or an untracked local `.env`. Never commit a token or other secret.

## DEPLOYMENT

### FadeHost with GitHub

1. Create a private GitHub repository and push the staged project files. `.gitignore` excludes `.env`, SQLite databases, exports, logs, backups, virtual environments, and user upload directories. `.env.example` contains only blank values.
2. Connect that repository to FadeHost and configure a Python background/worker service using Python 3.12.
3. Install dependencies with `pip install -r requirements.txt` and set the start command to `python bot.py`.
4. Configure `BOT_TOKEN`, `ADMIN_IDS`, and `DB_PATH` in FadeHost Secrets/Environment Variables. Never put `BOT_TOKEN` in GitHub, source files, or `.env.example`.
5. Use a persistent disk for the SQLite database. Set `DB_PATH=/data/finance_bot.db` only if `/data` is the mounted, writable directory. The application creates a missing parent directory when permissions allow. Without this setting, local operation continues to use `finance_bot.db` in the working directory.
6. The bot uses polling, so it needs a continuously running worker but no public port or webhook.

### Moving the existing database

Do not add `finance_bot.db` to GitHub. To move it later, stop the bot, copy the existing database file to the persistent volume without changing its contents, set `DB_PATH` to its new location, and then start the bot. Keep a separate untouched copy until the deployed data has been checked. Do not copy the database while the bot is running.

No automatic backup job is configured. Select and review a backup destination and retention policy before enabling production backups.

## Main files

- `bot.py` — bot handlers, SQLite initialization, reporting, and polling entry point.
- `requirements.txt` — pinned direct Python dependencies.
- `.python-version` — selected Python version.
- `.env.example` — blank local environment template.
