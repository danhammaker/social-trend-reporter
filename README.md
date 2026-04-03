# Daily Social Trend Scraper

This project builds a daily report of what was popular on social media the day before, with direct links to example memes and videos.

It uses public feeds that work without API keys:

- Reddit `top.json` for meme, viral-video, and explainer communities
- optional YouTube RSS feeds for channels you choose
- optional X trend + recent-search collection with an official bearer token
- optional TikTok public-page scraping from configured seed URLs

Each run:

1. pulls yesterday's items,
2. scores them,
3. groups them into a few topic clusters,
4. writes Markdown and HTML reports into `reports/`,
5. can email the formatted HTML version through Outlook, Gmail SMTP, or GitHub Actions secrets.

## Files

- `config/sources.json`: editable source list and output settings
- `scripts/Invoke-TrendReport.ps1`: main collector + summarizer
- `scripts/Register-DailyTrendTask.ps1`: optional Windows Task Scheduler setup
- `scripts/Save-ApiToken.ps1`: saves an encrypted X bearer token file for the current Windows user
- `.github/workflows/daily-social-trends.yml`: GitHub Actions schedule for unattended runs

## Run It

From the project root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Invoke-TrendReport.ps1
```

That writes a file like:

```text
reports/trend-report-2026-03-28.md
```

To test a specific date window, pass a reference date. The script always summarizes the day before that date:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Invoke-TrendReport.ps1 -ReportDate "2026-03-29"
```

To send the formatted HTML report by email with Outlook:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Invoke-TrendReport.ps1 -SendEmail -EmailTo "hammaker.dan@gmail.com"
```

To send by Gmail SMTP, first save an encrypted credential file, then run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Invoke-TrendReport.ps1 -SendEmail -EmailTo "hammaker.dan@gmail.com" -EmailMethod Smtp -SmtpCredentialPath ".\secrets\gmail-credential.json"
```

## Schedule It Daily

Create a Windows scheduled task that runs every day at 8:00 AM:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Register-DailyTrendTask.ps1
```

To choose a different time:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Register-DailyTrendTask.ps1 -RunAt "09:30"
```

## Run It In GitHub Actions

This repo includes a GitHub Actions workflow at `.github/workflows/daily-social-trends.yml` that can run on a schedule even when your local machine is asleep.

Set these repository secrets in GitHub:

- `SMTP_USERNAME`: your Gmail address
- `SMTP_PASSWORD`: your Gmail app password
- `X_BEARER_TOKEN`: your X bearer token, if you want X enabled

The workflow is scheduled in UTC but gates itself so it only runs at 7 AM New York time across daylight saving changes. You can also trigger it manually with `Run workflow` in the GitHub Actions tab.

## Customize Sources

Edit `config/sources.json` to add or remove:

- Reddit subreddits
- YouTube feeds, using either `feedUrl` or `channelId`
- X API settings such as WOEID, trend count, and bearer token path
- TikTok seed URLs to scrape for public video links
- output limits such as topic count and example count

## X Setup

Save your X bearer token into an encrypted file:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Save-ApiToken.ps1 -Token "YOUR_X_BEARER_TOKEN"
```

Then set `"x": { "enabled": true }` in `config/sources.json`.

## TikTok Setup

Set `"tiktok": { "enabled": true }` and add one or more public seed pages in `config/sources.json`, for example hashtag, discover, or account pages. The scraper will pull public TikTok video links from those pages and then extract metadata from the linked video pages.

## Notes

- This is feed-based scraping, so it is reliable and lightweight, but it does not use private platform APIs.
- Reddit scoring is based on post score plus comment count.
- YouTube RSS does not expose views or likes, so YouTube entries are included as concrete examples rather than hard popularity rankings.
- X support in this project uses the official API and is the most reliable way to add X content.
- TikTok support in this project is a best-effort public scraper. It is more fragile than the API-backed adapters and may need maintenance if TikTok changes its page structure.
- TikTok, Instagram, and other platforms can tighten rate limits or anti-bot protections over time.
- GitHub Actions schedules only run from the repository's default branch.
