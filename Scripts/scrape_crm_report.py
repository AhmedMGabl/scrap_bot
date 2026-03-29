# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from playwright.sync_api import sync_playwright
import requests
import time, os, sys, json
from datetime import datetime, timedelta

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

CRM_URL       = "https://crm.51talk.com/scReportForms/sc_call_info_new?userType=sc_group"
CRM_LOGIN_URL = "https://crm.51talk.com/admin/admin_login.php?login_employee_type=sideline&redirect_uri="
CRM_USERNAME  = "51Hany"
CRM_PASSWORD  = "b%7DWWtm"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

JS_EXTRACT = """
() => {
  const tables = document.querySelectorAll('table');
  const dataTable = Array.from(tables).find(t => t.textContent.includes('Total valid calls'));
  if (!dataTable) return JSON.stringify({error: 'no table'});
  const allRows = Array.from(dataTable.querySelectorAll('tr'));
  const headers = Array.from(allRows[0].querySelectorAll('th')).map(th => th.textContent.trim());
  const data = [];
  for (const row of allRows.slice(1)) {
    const cells = Array.from(row.querySelectorAll('td')).map(td => td.textContent.trim());
    if (cells.length < 10) continue;
    if (!cells[1] || cells[1] === '/') continue;
    // Skip group header rows and sub-total rows: agent rows always have a numeric serial
    if (!/^\d+$/.test(cells[0])) continue;
    data.push(cells);
  }
  return JSON.stringify({headers, data});
}
"""


def _try_requests(cookie_file, today_str, rawdata_file):
    try:
        import json as _json
        with open(cookie_file) as f:
            cookies = _json.load(f)
        resp = requests.post(
            CRM_URL,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Referer": CRM_URL,
            },
            cookies=cookies,
            data={
                "start_date": today_str, "end_date": today_str,
                "today_start_time": "00:00:00", "today_end_time": "23:59:59",
                "is_show_group": "y", "": "submit",
            },
            timeout=30
        )
        if resp.status_code != 200 or "Total valid calls" not in resp.text:
            print(f"  HTTP {resp.status_code}, falling back to browser...")
            return False
        try:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(resp.text, "html.parser")
        except ImportError:
            import html.parser
            print("  bs4 not available, falling back to browser...")
            return False
        tables = soup.find_all("table")
        data_table = next((t for t in tables if "Total valid calls" in t.get_text()), None)
        if not data_table:
            print("  Table not found in response, falling back to browser...")
            return False
        all_rows = data_table.find_all("tr")
        headers = [th.get_text(strip=True) for th in all_rows[0].find_all(["th", "td"])]
        rows = []
        for row in all_rows[1:]:
            cells = [td.get_text(strip=True) for td in row.find_all("td")]
            if len(cells) < 10 or not cells[1] or cells[1] == "/" or not cells[0].isdigit():
                continue
            rows.append(cells)
        if not rows:
            print("  No data rows in response (calls may be zero), saving zeros...")
        import pandas as pd
        df = pd.DataFrame(rows, columns=headers if rows and len(headers) == len(rows[0]) else None)
        col_map = {headers[0]: "Serial", headers[1]: "SC", headers[4]: "Total number of calls",
                   headers[5]: "Total valid calls", headers[13]: "Total effective call time/Minute",
                   headers[14]: "Average call time/Minute"} if headers else {}
        df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})
        for col in ["Total valid calls", "Total effective call time/Minute", "Average call time/Minute", "Total number of calls"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        print(f"  Shape: {df.shape}")
        with pd.ExcelWriter(rawdata_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="1", index=False)
        print(f"  Saved {len(df)} rows via requests")
        return True
    except Exception as ex:
        print(f"  Requests approach failed: {ex}")
        return False

def scrape_crm_report():
    script_dir   = os.path.dirname(os.path.abspath(__file__))
    parent_dir   = os.path.dirname(script_dir)
    rawdata_file = os.path.join(parent_dir, "Input", "rawdata.xlsx")
    PROFILE_DIR  = os.path.join(script_dir, "chrome_profile")
    now          = datetime.now()
    # After midnight but before noon: report on yesterday's shift
    if now.hour < 12:
        target_date = now - timedelta(days=1)
    else:
        target_date = now
    today_str    = target_date.strftime("%Y-%m-%d")

    print("=" * 60)
    print("CRM Call Report Scraper")
    print("=" * 60)
    print(f"Target date: {today_str}")

    cookie_file = os.path.join(script_dir, "crm_cookies.json")
    if os.path.exists(cookie_file):
        print("Step 1: Trying requests with saved cookies...")
        if _try_requests(cookie_file, today_str, rawdata_file):
            print("DONE")
            return
    else:
        print("Step 1: No cookie file, using browser...")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            executable_path=CHROME_PATH,
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(viewport={"width": 1400, "height": 900})
        page = context.new_page()
        page.set_default_timeout(60000)

        try:
            print("Step 1: Logging in to CRM...")
            page.goto(CRM_LOGIN_URL, wait_until="domcontentloaded", timeout=30000)
            if "admin_login" in page.url:
                page.locator("#user_name").fill(CRM_USERNAME)
                page.locator("#pwd").fill(CRM_PASSWORD)
                page.locator("#Submit").click()
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    pass  # domcontentloaded fired; continue regardless
                print(f"  After login: {page.url}")
            print("Step 2: Navigating to report page...")
            page.goto(CRM_URL, wait_until="domcontentloaded", timeout=60000)
            print(f"  Page: {page.url}")

            # Step 2: Set start_date to today via JS, set end_date to tomorrow via UI picker
            tomorrow      = target_date + timedelta(days=1)
            tomorrow_str  = tomorrow.strftime("%Y-%m-%d")
            tomorrow_day  = str(tomorrow.day)
            print(f"Step 2: Setting date range {today_str} to {tomorrow_str}...")
            try:
                page.evaluate(f"document.getElementById('start_date').value = '{today_str}'")
                time.sleep(0.3)
                print("  Start date: ok")
            except Exception as e:
                print(f"  WARNING: start_date set failed: {e}")

            # Click end_date input to open SelectDate() overlay picker, then select tomorrow
            page.click('input[name="end_date"]')
            page.wait_for_timeout(1200)
            # Calendar day cells have style="cursor: pointer" — target exactly those
            page.locator('td[style*="cursor"]').filter(has_text=tomorrow_day).first.click()
            page.wait_for_timeout(300)
            print(f"  End date picker: selected {tomorrow_str}")

            # Keep is_show_group checked (group view) to get ALL agents including zero-call agents
            # Individual view misses agents who have no calls on the queried date

            # Step 3: Click submit and wait for fresh results
            print("Step 3: Submitting query...")
            page.wait_for_timeout(500)  # let picker close fully before submitting
            submit = page.query_selector('input[type="submit"][value="submit"], input[value="submit"]')
            if submit:
                submit.click()
                print("  Submitted. Waiting for results to reload...")
                # Wait for the table to detach (page reloads) then reappear with fresh data
                try:
                    page.wait_for_selector("table:has-text('Total valid calls')", state="detached", timeout=10000)
                except Exception:
                    pass  # page may reload too fast to catch detach
                page.wait_for_selector("table:has-text('Total valid calls')", timeout=30000)
                page.wait_for_load_state("networkidle", timeout=30000)
                print("  Data table loaded.")
            else:
                print("  WARNING: submit button not found")
                time.sleep(3)

            # Step 5: Extract table data
            print("Step 4: Extracting table data...")
            result = page.evaluate(JS_EXTRACT)
            parsed = json.loads(result)

            if "error" in parsed:
                raise Exception(f"Table not found: {parsed['error']}")

            headers = parsed["headers"]
            rows    = parsed["data"]
            print(f"  Headers: {headers}")
            print(f"  Data rows: {len(rows)}")

            if not rows:
                print("  WARNING: No data rows found (possibly no calls today yet)")
                browser.close()
                return

            # Step 6: Build DataFrame with clean column names
            df = pd.DataFrame(rows, columns=headers if len(headers) == len(rows[0]) else None)

            # Rename key columns
            col_map = {
                headers[0]: 'Serial',
                headers[1]: 'SC',
                headers[4]: 'Total number of calls',
                headers[5]: 'Total valid calls',
                headers[13]: 'Total effective call time/Minute',
                headers[14]: 'Average call time/Minute',
            }
            df = df.rename(columns=col_map)

            # Convert numeric columns
            for col in ['Total valid calls', 'Total effective call time/Minute', 'Average call time/Minute', 'Total number of calls']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            print(f"  Shape: {df.shape}")
            print(df[['SC','Total valid calls','Total effective call time/Minute','Average call time/Minute']].head(5).to_string())

            # Step 7: Save to rawdata.xlsx tab 1
            print("Step 5: Saving to rawdata.xlsx sheet '1'...")
            with pd.ExcelWriter(rawdata_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name="1", index=False)
            print(f"  Saved {len(df)} rows to sheet '1' in {rawdata_file}")

            # Auto-refresh cookies for next run
            try:
                fresh = context.cookies()
                cookie_dict = {c["name"]: c["value"] for c in fresh}
                with open(cookie_file, "w") as cf:
                    json.dump(cookie_dict, cf)
                print("  Cookies refreshed.")
            except Exception as ce:
                print(f"  Cookie save warning: {ce}")

        except Exception as e:
            print(f"ERROR: {e}")
            import traceback
            traceback.print_exc()
        finally:
            browser.close()
            print("DONE")

if __name__ == "__main__":
    scrape_crm_report()
