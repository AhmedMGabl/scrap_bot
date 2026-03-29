# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose
Automated daily pipeline that scrapes CRM and AMS data, generates CM and EA agent performance dashboards (HTML + PNG), and sends them to DingTalk (and optionally Lark).

## How to Run

```bash
# Run the full pipeline (scrape -> generate -> send)
python run_daily_report.py

# Set up Windows Task Scheduler (run once as Administrator)
python setup_schedule.py

# Remove all scheduled tasks
python remove_schedule.py
```

The pipeline skips Fridays automatically. AMS scraping opens a browser window -- log in manually if auto-login fails (script waits 120s).

## Folder Structure

```
scrap_bot/
├── run_daily_report.py          # Entry point -- orchestrates all phases
├── generate_cm_report.py        # CM data reading, merging, HTML/PNG generation
├── generate_ea_report.py        # EA data reading, aggregation (imports CM's HTML funcs)
├── setup_schedule.py            # Windows Task Scheduler config
├── Input/
│   ├── rawdata.xlsx             # Sheet 1: CRM data; Sheet 2: IUR/AMS data
│   ├── Team Structure.xlsx      # Sheet CM: Team+CRM cols; Sheet EA: Chinese cols renamed
│   └── EA_rawdata_Nov_Jan.xlsx  # Historical multi-month EA data
├── Scripts/
│   ├── scrape_crm_report.py     # Playwright scraper for crm.51talk.com
│   ├── scrape_iur_report.py     # Playwright scraper for ams.51talkjr.com
│   └── html_report_generator.py # generate_html_team_report() (Team Summary view)
├── Output/
│   ├── CM_*.html / CM_*.png     # Generated CM dashboards
│   ├── EA_*.html / EA_*.png     # Generated EA dashboards
│   └── logs/YYYY-MM-DD_HH-MM.log
└── image_host/                  # Local clone of GitHub image repo (for DingTalk)
```

## Pipeline Phases

| Phase | File | What it does |
|-------|------|-------------|
| 1 - CRM Scrape | Scripts/scrape_crm_report.py | Logs into crm.51talk.com, exports today call data to rawdata.xlsx sheet 1 |
| 2 - AMS Scrape | Scripts/scrape_iur_report.py | Logs into ams.51talkjr.com, exports IUR data to rawdata.xlsx sheet 2 |
| 3 - CM Dashboard | generate_cm_report.py + Scripts/html_report_generator.py | Merges CRM + IUR + Team Structure, generates 4 CM HTML reports + PNGs |
| 4 - EA Dashboard | generate_ea_report.py | Reads EA historical data + Team Structure, generates 4 EA HTML reports + PNGs |
| 5 - Lark Cards | run_daily_report.py (run_send_cards) | Currently commented out -- uploads PNGs and sends interactive cards to Lark |
| 6 - DingTalk | run_daily_report.py (run_send_dingtalk) | Active -- pushes PNGs to GitHub image host, sends markdown card via webhook |

## Current Active Config (run_daily_report.py)

- TEST_MODE = True -- sends to Hany testing group, not Maze Runners production group
- Lark card sending (run_send_cards) is commented out; DingTalk (run_send_dingtalk) is active
- To switch to production: set TEST_MODE = False (sends to LARK_CHAT_ID_PROD)

## Credentials & Config (hardcoded in run_daily_report.py)
- CRM: username 51Hany, password b%7DWWtm -- crm.51talk.com
- AMS: username 51hany, password Hyoussef@51 -- ams.51talkjr.com
- Lark App ID: cli_a9bf7d0d8438dbdc
- Lark Chat IDs: oc_cc12fe7005d8a9fa8b8eb51e9193eeec (Maze Runners/prod), oc_1ab849cf11a8505ae909eff1928cd052 (Hany/test)
- DingTalk webhook token in DINGTALK_WEBHOOK_URL

## sys.path Import Order (CRITICAL)

run_daily_report.py inserts paths in this order:
```python
sys.path.insert(0, os.path.join(SCRIPT_DIR, "Scripts"))
sys.path.insert(0, SCRIPT_DIR)  # ROOT wins -- index 0 overrides Scripts
```
The ROOT generate_cm_report.py must win over Scripts/generate_cm_report.py because only
the root version has read_duration_data, read_iur_data, read_cm_structure, merge_all_data.
generate_ea_report.py imports these from generate_cm_report and relies on the same resolution.

## Key Data Columns

### rawdata.xlsx sheet 1 (CRM)
Headers at row 1, no skiprows needed: Serial, SC, Total valid calls,
Total effective call time/Minute, Average call time/Minute

### rawdata.xlsx sheet 2 (IUR/AMS)
Columns: org_name1, useraccount1, Class completed (and others).
The useraccount1 column maps to agent CRM usernames.

### Team Structure.xlsx
- Sheet CM: columns Team, CRM (agent username)
- Sheet EA: Chinese column names, renamed to Team, CRM

## Report Logic -- Avg Call Time/Min
Most important metric. Each value is int(round(agent_value)).

- Per agent: taken directly from CRM column "Average call time/Minute"
- Team TOTAL row: sum(int(round(x)) for each agent) -- NOT duration/calls
- Team AVERAGE row: mean(int(round(x)) for each agent)
- Team Summary table per-team value = same formula as Separate Teams TOTAL

Implemented in:
- generate_cm_report.py -> generate_html_separate_teams_report() around line 550-568
- Scripts/html_report_generator.py -> generate_html_team_report() around line 18-22

## Color Coding (Excel-style 3-color scale)
Applied to Total Duration (Min) and Avg Call Time/Min columns. Min/max excludes zero-data teams.

- bg-very-low  #F8696B  bottom 17%
- bg-low       #FCAA75  17-33%
- bg-medium-low #FFEB84 33-50%
- bg-medium    #C6E5B5  50-67%
- bg-medium-high #9FD899 67-83%
- bg-high      #63BE7B  top 17%

TOTAL row = blue (#dbeafe, border #3b82f6)
AVERAGE row = purple (#f3e8ff, border #a855f7)
Individual report = NO total/average rows

## EA vs CM Differences
- CM data source: CRM scraped daily; EA data source: EA_rawdata_Nov_Jan.xlsx (historical multi-month)
- CM Avg Call Time/Min: taken directly from CRM column; EA: mean of monthly averages
- CM Classes Completed: from IUR (AMS scrape); EA: always 0 (not tracked)

## Known Fixed Bugs -- Do Not Reintroduce
1. read_duration_data: no skiprows (scraper saves headers at row 1)
2. read_iur_data: column is useraccount1, NOT org_name1/useraccount1
3. scrape_and_update_rawdata() in root generate_cm_report.py: dead code, had double-except bug, now fixed to nested try-except
4. EA Excel team summary Avg Call Time/Min: uses sum of individual rounded values, not Duration/Calls
5. Avg Eff. Calls: int(round(calls/members)) per agent; TOTAL row uses int(round(mean of team values))
6. CRM scraper: do NOT uncheck is_show_group -- group view shows ALL agents incl. zero-call; individual view misses zero-call agents. JS_EXTRACT and _try_requests filter agent rows by numeric serial (cells[0].isdigit()) to skip group headers and sub-totals
7. Color scale in Team Summary and Separate Teams: exclude zero-data teams from min/max so non-zero teams get accurate color gradient
8. Separate Teams TOTAL row Avg Call Time/Min: use Duration/Members, NOT Duration/Eff.Calls
9. Team Summary TOTAL row Avg Call Time/Min: use int(round(total_duration/total_members))

## Dependencies
```bash
pip install pandas openpyxl requests playwright psutil
playwright install chromium
```

## Platform
Windows 11, Python 3.11, bash shell (use Unix paths in scripts).
Chrome at: C:\Program Files\Google\Chrome\Application\chrome.exe

## Scheduler (Windows Task Scheduler)
Schedule: Morning (12:30, 13:30, 14:30, 15:30) -- Night (22:00, 23:00, 00:00) -- Saturday overflow (01:00 Sunday).
Each run logs to Output/logs/YYYY-MM-DD_HH-MM.log.
