# scrap_bot — Daily Duration Report Pipeline

## Project Purpose
Automated daily pipeline that scrapes CRM and AMS data, generates CM and EA agent
performance dashboards (HTML + PNG), and sends them to a Lark group chat.

## How to Run
```bash
cd "C:/Users/high tech/Documents/GitHub/scrap_bot"
python run_daily_report.py
```
Runs all 5 phases. AMS scrape opens a browser — log in manually if auto-login fails,
the script waits up to 120s.

## Folder Structure
```
scrap_bot/
├── run_daily_report.py          <- MAIN ENTRY POINT
├── generate_cm_report.py        <- CM data functions + HTML generators
├── generate_ea_report.py        <- EA data functions + Excel generators
├── Scripts/
│   ├── scrape_crm_report.py     <- Playwright CRM scraper (crm.51talk.com)
│   ├── scrape_iur_report.py     <- Playwright AMS scraper (ams.51talkjr.com)
│   ├── html_report_generator.py <- Team Summary HTML generator
│   ├── crm_cookies.json         <- Saved CRM login cookies
│   ├── chrome_profile/          <- Persistent Playwright Chrome profile
│   └── downloads/               <- IUR Excel files downloaded from AMS
├── Input/
│   ├── rawdata.xlsx             <- Sheet '1': CRM | Sheet '2': IUR
│   ├── Team Structure.xlsx      <- Sheet 'CM': CM teams | Sheet 'EA': EA teams
│   └── EA_rawdata_Nov_Jan.xlsx  <- EA historical data (Nov-Jan)
└── Output/
    ├── CM_Individual_Report.html/png
    ├── CM_Team_Summary.html/png
    ├── CM_Separate_Teams.html/png
    ├── CM_Bottom20.html/png
    ├── EA_Individual_Report.html/png
    ├── EA_Team_Summary.html/png
    ├── EA_Separate_Teams.html/png
    └── EA_Bottom20.html/png
```

## Pipeline Phases
| Phase | Description |
|-------|-------------|
| 1 - CRM Scrape  | Login crm.51talk.com, export today call data -> rawdata.xlsx sheet 1 |
| 2 - AMS Scrape  | Login ams.51talkjr.com, export IUR data -> rawdata.xlsx sheet 2 |
| 3 - CM Dashboard | Merge CRM + IUR + Team Structure -> 4 CM HTML reports + PNGs |
| 4 - EA Dashboard | Read EA historical + Team Structure -> 4 EA HTML reports + PNGs |
| 5 - Lark Cards  | Upload PNGs, send interactive cards to Lark group |

## Credentials (hardcoded in run_daily_report.py)
- CRM login: 51Hany / b%7DWWtm (crm.51talk.com)
- AMS login: 51hany / Hyoussef@51 (ams.51talkjr.com)
- Lark App ID: cli_a9bf7d0d8438dbdc
- Lark Chat ID: oc_1ab849cf11a8505ae909eff1928cd052

## sys.path Import Order (CRITICAL — do not change)
```python
sys.path.insert(0, os.path.join(SCRIPT_DIR, "Scripts"))  # added first
sys.path.insert(0, SCRIPT_DIR)                            # root = highest priority
```
ROOT generate_cm_report.py must win over Scripts/generate_cm_report.py.
Only the root version has: read_duration_data, read_iur_data, read_cm_structure, merge_all_data.

## Key Data Columns

### rawdata.xlsx sheet 1 (CRM) — headers at row 1, no skiprows
Columns used: Serial, SC, Total valid calls,
              Total effective call time/Minute, Average call time/Minute

### rawdata.xlsx sheet 2 (IUR/AMS)
Columns: org_name1, useraccount1, Class completed
useraccount1 = agent CRM username (used for matching)

### Team Structure.xlsx
- Sheet 'CM': Team, CRM
- Sheet 'EA': Chinese headers, renamed to Team, CRM

## Avg Call Time/Min Formula (IMPORTANT)
- Per agent: int(round(value)) from CRM "Average call time/Minute"
- Team TOTAL row: sum of int(round(x)) for each agent — NOT duration/calls
- Team AVERAGE row: mean of int(round(x)) for each agent
- Team Summary table: same formula as Separate Teams TOTAL row per team

## Color Coding (Excel 3-color scale)
Columns: Total Duration (Min) and Avg Call Time/Min
- bg-very-low  #F8696B  bottom 17% (red)
- bg-low       #FCAA75  17-33%
- bg-medium-low #FFEB84 33-50%
- bg-medium    #C6E5B5  50-67%
- bg-medium-high #9FD899 67-83%
- bg-high      #63BE7B  top 17% (green)

TOTAL row   = blue (#dbeafe, border #3b82f6)
AVERAGE row = purple (#f3e8ff, border #a855f7)
Individual report = NO total/average rows

## EA vs CM
- CM: scraped daily from CRM, Classes Completed from IUR
- EA: historical file (EA_rawdata_Nov_Jan.xlsx), Classes Completed = 0
- EA Avg Call Time/Min: mean of monthly averages across months

## Fixed Bugs — Do Not Reintroduce
1. read_duration_data: NO skiprows (scraper saves with headers at row 1)
2. read_iur_data: column = 'useraccount1', NOT 'org_name1/useraccount1'
3. scrape_and_update_rawdata(): dead code, had double-except — fixed to nested try-except
4. EA Excel Avg Call Time/Min: uses sum of individual rounded values, not Duration/Calls
5. Avg Eff. Calls: 1 decimal — round(calls/members, 1)

## Dependencies
```
pip install pandas openpyxl playwright requests
playwright install chromium
```

## Platform
Windows 11, Python 3.11
Chrome: C:\Program Files\Google\Chrome\Application\chrome.exe
