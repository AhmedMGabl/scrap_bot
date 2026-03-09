# scrap_bot — Daily Duration Report Pipeline

## Project Purpose
Automated daily pipeline that scrapes CRM and AMS data, generates CM and EA agent performance dashboards (HTML + PNG), and sends them to a Lark group chat.

## How to Run
\============================================================
PHASE 1: Scraping CRM (crm.51talk.com)
============================================================
============================================================
CRM Call Report Scraper
============================================================
Target date: 2026-03-09
Step 1: Trying requests with saved cookies...
  bs4 not available, falling back to browser...
Step 1: Logging in to CRM...
  After login: https://crm.51talk.com/admin/login.php
Step 2: Navigating to report page...
  Page: https://crm.51talk.com/scReportForms/sc_call_info_new?userType=sc_group
Step 2: Setting date to 2026-03-09...
Step 3: Submitting query...
  Submitted. Waiting for data...
Step 4: Extracting table data...
  Headers: ['Serial number', 'SC', 'First call', 'Last call', 'Total number of calls', 'Total valid calls', '<1Number of minute(s)', '<1Number of minute(s)/Number of effective call times%', '1-3Number of minute(s)', '1-3Number of minute(s)/Number of effective call times%', '>3Number of minute(s)', '>3Number of minute(s)/Number of effective call times%', 'Rate of effective calls', 'Total effective call time/Minute', 'Average call time/Minute']
  Data rows: 645
  Shape: (645, 15)
                     SC  Total valid calls  Total effective call time/Minute  Average call time/Minute
0                51Hany                0.0                               0.0                       0.0
1               stcc001                0.0                               0.0                       0.0
2  51abdulrahma-testref                0.0                               0.0                       0.0
3     51abualhaija-test                0.0                               0.0                       0.0
4    51albisani-testref                0.0                               0.0                       0.0
Step 5: Saving to rawdata.xlsx sheet '1'...
  Saved 645 rows to sheet '1' in D:\Daily reports\Daily Duration Report\Input
awdata.xlsx
DONE
OK CRM scraping done.
============================================================
PHASE 2: Scraping AMS (ams.51talkjr.com)
============================================================
============================================================
IUR New Report Scraper
============================================================
Step 1: Logging in to AMS...
  Auto-login failed (Timeout 15000ms exceeded.
=========================== logs ===========================
waiting for navigation to **/login-turn** until 'load'
  navigated to https://ams.51talkjr.com/#/login
  navigated to https://lp.51talkjr.com/#/welcome
  navigated to https://lp.51talkjr.com/#/welcome
  navigated to https://lp.51talkjr.com/#/welcome
  navigated to https://lp.51talkjr.com/#/welcome
============================================================), log in manually
Step 2: Waiting up to 120s for LP BI frame...
  LP welcome, navigating to IUR...
  LP at s://lp.51talkjr.com/#/data-center/business/iur_new, 2 frames...
  BI frame found! tabs=['Total', 'Details']
Step 3: Clicking Total tab...
  Total tab clicked
Step 4: Setting date to today and querying...
  Start date: ok
  End date: ok
  查询 clicked, waiting 10s...
Step 5: Exporting report...
  Grid/expand li: clicked
  Download icon: clicked
  Confirm: clicked:BUTTON:确 定
  Downloaded: D:\Daily reports\Daily Duration Report\Scripts\downloads\海外IUR_NEW_DATA_20260309_1448.xlsx
Step 6: Saving to rawdata.xlsx sheet 2...
  Shape: 220 rows x 17 cols
  Columns: ['org_name1', 'useraccount1', 'Class opened', 'Class completed', 'Total students', 'Booked students', 'Booked rate', 'Attended students', 'Attended rate', 'Coverage rate', 'Total classes', 'Booked classes', 'Class Booked rate', 'Attended classes', 'Class Attended rate', 'Absent classes', 'Class Coverage rate']
         org_name1    useraccount1  Class opened  Class completed  Total students  Booked students  Booked rate  Attended students  Attended rate  Coverage rate  Total classes  Booked classes  Class Booked rate  Attended classes  Class Attended rate  Absent classes  Class Coverage rate
0  ME-EGLP-GCC01小组   51abdelrahman             0                0               4                0     0.000000                  0            NaN       0.000000              4               0           0.000000                 0                  NaN               0             0.000000
1  ME-EGLP-GCC01小组      EGLP-OmarM             5                1              57               33     0.578947                 17       0.515152       0.298246             85              34           0.400000                17                 0.50               7             0.200000
2  ME-EGLP-GCC01小组  EGLP-ahmedalfy             0                0              38               24     0.631579                 13       0.541667       0.342105             51              25           0.490196                13                 0.52              11             0.254902
  Saved to sheet 2 in D:\Daily reports\Daily Duration Report\Input
awdata.xlsx
DONE
OK AMS scraping done.
============================================================
PHASE 3: Generating CM Dashboard
============================================================
Columns found: ['Serial', 'SC', 'First call', 'Last call', 'Total number of calls']...
HTML individual report saved to: D:\Daily reports\Daily Duration Report\Output\CM_Individual_Report.html
HTML report saved to: D:\Daily reports\Daily Duration Report\Output\CM_Team_Summary.html
HTML separate teams report saved to: D:\Daily reports\Daily Duration Report\Output\CM_Separate_Teams.html
HTML bottom 20 report saved to: D:\Daily reports\Daily Duration Report\Output\CM_Bottom20.html

Generating screenshots...
Screenshot saved: D:\Daily reports\Daily Duration Report\Output\CM_Individual_Report.png
Screenshot saved: D:\Daily reports\Daily Duration Report\Output\CM_Team_Summary.png
Screenshot saved: D:\Daily reports\Daily Duration Report\Output\CM_Separate_Teams.png
Screenshot saved: D:\Daily reports\Daily Duration Report\Output\CM_Bottom20.png
All screenshots generated successfully!
OK CM dashboard done.
============================================================
PHASE 5: Sending cards to Lark group
============================================================
  Token obtained
  Uploading CM images...
  CM card sent
OK Cards sent.This runs all 5 phases sequentially. The AMS scrape opens a browser window — log in manually if auto-login fails, the script waits up to 120s.

## Folder Structure
## Pipeline Phases
| Phase | What it does |
|-------|-------------|
| 1 - CRM Scrape | Logs into crm.51talk.com, exports today call data to rawdata.xlsx sheet 1 |
| 2 - AMS Scrape | Logs into ams.51talkjr.com, exports IUR data to rawdata.xlsx sheet 2 |
| 3 - CM Dashboard | Merges CRM + IUR + Team Structure, generates 4 CM HTML reports + PNGs |
| 4 - EA Dashboard | Reads EA historical data + Team Structure, generates 4 EA HTML reports + PNGs |
| 5 - Lark Cards | Uploads PNGs and sends interactive cards to Lark group |

## Credentials & Config (hardcoded in run_daily_report.py)
- CRM: username 51Hany, password b%7DWWtm — crm.51talk.com
- AMS: username 51hany, password Hyoussef@51 — ams.51talkjr.com
- Lark App ID: cli_a9bf7d0d8438dbdc
- Lark Chat ID: oc_1ab849cf11a8505ae909eff1928cd052

## sys.path Import Order (CRITICAL)
run_daily_report.py inserts paths so the ROOT files take priority:
\The ROOT generate_cm_report.py must win over Scripts/generate_cm_report.py because only
the root version has read_duration_data, read_iur_data, read_cm_structure, merge_all_data.

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

## Report Logic — Avg Call Time/Min
Most important metric. Each value is int(round(agent_value)).

- Per agent: taken directly from CRM column "Average call time/Minute"
- Team TOTAL row: sum(int(round(x)) for each agent) — NOT duration/calls
- Team AVERAGE row: mean(int(round(x)) for each agent)
- Team Summary table per-team value = same formula as Separate Teams TOTAL

Implemented in:
- generate_cm_report.py -> generate_html_separate_teams_report() around line 550-568
- Scripts/html_report_generator.py -> generate_html_team_report() around line 18-22

## Color Coding (Excel-style 3-color scale)
Applied to Total Duration (Min) and Avg Call Time/Min columns:
- bg-very-low  #F8696B  bottom 17%
- bg-low       #FCAA75  17-33%
- bg-medium-low #FFEB84 33-50%
- bg-medium    #C6E5B5  50-67%
- bg-medium-high #9FD899 67-83%
- bg-high      #63BE7B  top 17%

TOTAL row  = blue  (#dbeafe, border #3b82f6)
AVERAGE row = purple (#f3e8ff, border #a855f7)
Individual report = NO total/average rows

## EA vs CM Differences
- CM data source: CRM scraped daily
- EA data source: EA_rawdata_Nov_Jan.xlsx (historical multi-month)
- CM Avg Call Time/Min: taken directly from CRM column
- EA Avg Call Time/Min: mean of monthly averages (aggregated across months)
- CM Classes Completed: from IUR (AMS scrape)
- EA Classes Completed: always 0 (not tracked for EA)

## Known Fixed Bugs — Do Not Reintroduce
1. read_duration_data: no skiprows (scraper saves headers at row 1)
2. read_iur_data: column is useraccount1, NOT org_name1/useraccount1
3. scrape_and_update_rawdata() in root generate_cm_report.py: dead code, had double-except bug, now fixed to nested try-except
4. EA Excel team summary Avg Call Time/Min: uses sum of individual rounded values, not Duration/Calls
5. Avg Eff. Calls: always 1 decimal (e.g. 0.4), computed as round(calls/members, 1)

## Dependencies
\Install Playwright browsers: playwright install chromium

## Platform
Windows 11, Python 3.11, bash shell (use Unix paths)
Chrome at: C:\Program Files\Google\Chrome\Application\chrome.exe

## Scheduler (Windows Task Scheduler)

Run once as Administrator:
```bash
python setup_schedule.py
```
To remove all tasks: `python remove_schedule.py`

Schedule:
- Morning every day: 12:30, 13:30, 14:30, 15:30
- Night every day:   22:00, 23:00, 00:00
- Saturday overflow: 01:00 Sunday (Saturday night extended shift)

Each run logs to: `Output/logs/YYYY-MM-DD_HH-MM.log`
