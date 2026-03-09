# Design: Scheduler + Performance Improvements

**Date:** 2026-03-09

## Goals
1. Automatically run the daily report pipeline on a fixed hourly schedule
2. Reduce per-run execution time via smart waits, cookie persistence, and parallel uploads

---

## 1. Scheduler

**Approach:** Windows Task Scheduler via a one-time setup script.

### Files
- `setup_schedule.py` — creates all tasks under a `scrap_bot` folder in Task Scheduler (must run as Administrator)
- `remove_schedule.py` — deletes all scrap_bot tasks

### Schedule

| Task name | Times | Days |
|-----------|-------|------|
| `scrap_bot_morning` | 12:30, 13:30, 14:30, 15:30 | Every day |
| `scrap_bot_night` | 22:00, 23:00, 00:00 | Every day |
| `scrap_bot_saturday_late` | 01:00 | Sunday only |

**Saturday total runs:** 12:30 → 13:30 → 14:30 → 15:30 → 22:00 → 23:00 → 00:00 → 01:00

Each task runs:
```
python "C:/Users/high tech/Documents/GitHub/scrap_bot/run_daily_report.py"
```
Output is logged to `Output/logs/YYYY-MM-DD_HH-MM.log`.

---

## 2. Performance Improvements

### 2a. Replace hardcoded sleeps with smart waits

| File | Current | Replace with |
|------|---------|--------------|
| `scrape_crm_report.py` | `time.sleep(8)` after submit | `page.wait_for_selector('table:has-text("Total valid calls")')` |
| `scrape_iur_report.py` | `time.sleep(10)` after query | `page.wait_for_selector()` on results table |
| `scrape_iur_report.py` | `time.sleep(2)` / `time.sleep(5)` | `page.wait_for_load_state('networkidle')` |
| `generate_cm_report.py` | `time.sleep(0.5)` per screenshot | `page.wait_for_load_state('networkidle')` |

**Estimated savings:** 15-25 seconds per run.

### 2b. Auto-refresh CRM cookies

After every successful Playwright browser login in `scrape_crm_report.py`, save fresh cookies to `crm_cookies.json`. Currently cookies are never auto-updated, so once they expire every run falls back to the slow browser path (~15-20 min). Auto-save restores the fast HTTP path after the next browser login.

### 2c. Parallel Lark image uploads

In `run_daily_report.py`, replace sequential image uploads with `concurrent.futures.ThreadPoolExecutor`. Both images are uploaded simultaneously, then assembled into a single card message as before.

**Estimated savings:** ~30 seconds per run.

### 2d. Run logging

`run_daily_report.py` redirects stdout to `Output/logs/YYYY-MM-DD_HH-MM.log` at startup so every scheduled run is recorded for debugging.

---

## Summary of Files Changed

| File | Change |
|------|--------|
| `setup_schedule.py` | New — registers Task Scheduler tasks |
| `remove_schedule.py` | New — removes Task Scheduler tasks |
| `run_daily_report.py` | Parallel uploads + log file setup |
| `Scripts/scrape_crm_report.py` | Smart waits + auto-save cookies |
| `Scripts/scrape_iur_report.py` | Smart waits |
| `generate_cm_report.py` | Smart waits in generate_screenshots |
