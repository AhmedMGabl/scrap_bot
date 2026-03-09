# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from playwright.sync_api import sync_playwright
import time, os, sys, glob
from datetime import datetime

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

AMS_USERNAME  = "51hany"
AMS_PASSWORD  = "Hyoussef@51"
AMS_LOGIN_URL = "https://ams.51talkjr.com/#/login"
LP_IUR_URL    = "https://lp.51talkjr.com/#/data-center/business/iur_new"


def scrape_iur_new_report():
    script_dir   = os.path.dirname(os.path.abspath(__file__))
    parent_dir   = os.path.dirname(script_dir)
    rawdata_file = os.path.join(parent_dir, "Input", "rawdata.xlsx")
    CHROME_PATH  = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    PROFILE_DIR  = os.path.join(script_dir, "chrome_profile")
    DOWNLOAD_DIR = os.path.join(script_dir, "downloads")
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    print("=" * 60)
    print("IUR New Report Scraper")
    print("=" * 60)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            executable_path=CHROME_PATH,
            headless=False,
            viewport={"width": 1400, "height": 900},
            args=["--disable-blink-features=AutomationControlled"],
            accept_downloads=True,
            downloads_path=DOWNLOAD_DIR,
        )
        page     = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(60000)
        lp_page  = None
        bi_frame = None

        try:
            # Step 1: Login
            print("Step 1: Logging in to AMS...")
            lp_page = context.pages[0] if context.pages else context.new_page()
            lp_page.goto(AMS_LOGIN_URL)
            try:
                lp_page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                lp_page.wait_for_timeout(2000)
            if "login" in lp_page.url and "login-turn" not in lp_page.url:
                try:
                    lp_page.wait_for_selector('input[placeholder*="手机号"]', timeout=8000)
                    lp_page.fill('input[placeholder*="手机号"]', AMS_USERNAME)
                    lp_page.fill('input[placeholder*="密码"]', AMS_PASSWORD)
                    lp_page.click('button:has-text("登录")')
                    lp_page.wait_for_url("**/login-turn**", timeout=15000)
                    print("  Login OK")
                except Exception as e:
                    print(f"  Auto-login failed ({e}), log in manually")

            # Step 2: Wait for LP BI frame
            print("Step 2: Waiting up to 120s for LP BI frame...")
            deadline   = time.time() + 120
            grid_found = False
            while time.time() < deadline:
                lp_candidate = None
                for pg in context.pages:
                    if "lp.51talkjr.com" in pg.url:
                        lp_candidate = pg
                        break
                if lp_candidate:
                    cur_url = lp_candidate.url
                    if "welcome" in cur_url or cur_url.endswith("#/"):
                        print("  LP welcome, navigating to IUR...")
                        lp_candidate.goto(LP_IUR_URL)
                        try:
                            lp_candidate.wait_for_load_state("networkidle", timeout=15000)
                        except Exception:
                            lp_candidate.wait_for_timeout(3000)
                    else:
                        for frm in lp_candidate.frames:
                            try:
                                tabs = frm.query_selector_all(".bi-tab-item-text")
                                if tabs:
                                    bi_frame   = frm
                                    lp_page    = lp_candidate
                                    grid_found = True
                                    print(f"  BI frame found! tabs={[t.text_content() for t in tabs]}")
                                    break
                            except Exception:
                                pass
                        if grid_found:
                            break
                        print(f"  LP at {cur_url[-50:]}, {len(lp_candidate.frames)} frames...")
                else:
                    cur_url = lp_page.url if lp_page else ""
                    if "login-turn" in cur_url:
                        print("  >> Please click the WhatsApp option <<")
                    else:
                        print(f"  Waiting... ({cur_url[-60:]})")
                time.sleep(2)
            if not grid_found:
                raise Exception("LP BI grid not found after 120s.")

            # Step 3: Click Total tab
            print("Step 3: Clicking Total tab...")
            total_tab = bi_frame.query_selector(".bi-tab-item-text[title='Total']")
            if total_tab:
                total_tab.click()
                print("  Total tab clicked")
                time.sleep(3)
            else:
                print("  WARNING: Total tab not found")

            # Step 4: Set date to today and query
            print("Step 4: Setting date to today and querying...")
            today_str = datetime.now().strftime("%Y-%m-%d")
            all_frames = [lp_page] + list(lp_page.frames)

            # Find and click the date input
            date_frame = None
            for sf in all_frames:
                try:
                    if sf.query_selector('input[placeholder="请选择时间"]'):
                        date_frame = sf
                        break
                except Exception:
                    pass

            if date_frame:
                try:
                    inp = date_frame.query_selector('input[placeholder="请选择时间"]')
                    inp.click()
                    lp_page.wait_for_selector("td[title='" + today_str + "']", timeout=5000)
                    # Click today cell in the calendar (both start and end)
                    for attempt in range(2):
                        sel = 'td[title="' + today_str + '"] .ant-picker-cell-inner'
                        js = "() => { var cells = document.querySelectorAll('" + sel + "'); if (cells.length) { cells[0].click(); return true; } return false; }"
                        clicked = date_frame.evaluate(js)
                        if not clicked:
                            clicked = lp_page.evaluate(js)
                        label = "Start" if attempt == 0 else "End"
                        print(f"  {label} date: {'ok' if clicked else 'not found'}")
                        lp_page.wait_for_timeout(300)
                    lp_page.keyboard.press("Escape")
                except Exception as e:
                    print(f"  WARNING: date range: {e}")
            else:
                print("  Date input not found, using existing filter")

            # Click query button
            query_btn = None
            for sf in all_frames:
                try:
                    for el in sf.query_selector_all("span, button"):
                        try:
                            txt = (el.text_content() or "").strip().replace(" ", "").replace("　", "")
                            if txt == "查询":
                                query_btn = el
                                break
                        except Exception:
                            pass
                    if query_btn:
                        break
                except Exception:
                    pass
            if query_btn:
                query_btn.click()
                print("  查询 clicked, waiting for results...")
                try:
                    lp_page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    lp_page.wait_for_timeout(3000)
            else:
                print("  WARNING: 查询 not found")
                time.sleep(5)




            # Step 5: Export - hover to reveal mini menu, click appstore expand, then download icon
            print("Step 5: Exporting report...")

            # Hover over the BI widget area to reveal the mini menu icons
            try:
                hover_target = bi_frame.query_selector('.dashboard-chart, .bi-design-render-table, .table-widget-component, .root-container')
                if hover_target:
                    hover_target.hover()
                    lp_page.wait_for_timeout(500)
                else:
                    # hover at center of frame
                    bi_frame.evaluate("() => { document.body.dispatchEvent(new MouseEvent('mouseover',{bubbles:true})); }")
                    lp_page.wait_for_timeout(500)
            except Exception as e:
                print(f"  Hover warning: {e}")

            lp_page.screenshot(path=os.path.join(parent_dir, "Output", "iur_before_export.png"))

            # 5a: Click the appstore/grid li (preview-mini-menu expand icon) to open download options
            clicked_grid = bi_frame.evaluate(
                "() => {"
                "  var el = document.querySelector('li.menu-expand-icon, li.advanced-tooltip, li.preview-mini-menu-list-item.menu-expand-icon');"
                "  if (!el) {"
                "    var lis = Array.from(document.querySelectorAll('li.preview-mini-menu-list-item'));"
                "    el = lis[lis.length-1];"
                "  }"
                "  if (el) { el.click(); return 'clicked'; } return 'not found';"
                "}"
            )
            print(f"  Grid/expand li: {clicked_grid}")
            try:
                bi_frame.wait_for_selector("svg.common-download-outlined-svg", timeout=3000)
            except Exception:
                lp_page.wait_for_timeout(1000)
            lp_page.screenshot(path=os.path.join(parent_dir, "Output", "iur_after_grid.png"))

            # 5b: Click the download icon (common-download-outlined-svg)
            JS_DL_ICON = (
                "() => {"
                "  var svg = document.querySelector('svg.common-download-outlined-svg, .common-download-outlined-svg');"
                "  if (!svg) return 'not found';"
                "  var btn = svg.closest('button') || svg.closest('li') || svg.closest('span') || svg.parentElement;"
                "  (btn || svg).click(); return 'clicked';"
                "}"
            )
            clicked_dl = bi_frame.evaluate(JS_DL_ICON)
            if clicked_dl == 'not found':
                clicked_dl = lp_page.evaluate(JS_DL_ICON)
            print(f"  Download icon: {clicked_dl}")
            try:
                lp_page.wait_for_selector("button:has-text('确定')", timeout=3000)
            except Exception:
                lp_page.wait_for_timeout(800)
            lp_page.screenshot(path=os.path.join(parent_dir, "Output", "iur_after_export_menu.png"))

            # 5c: Click confirm (确 定) and capture download
            JS_CONFIRM = (
                "() => {"
                "  var btns = Array.from(document.querySelectorAll('button, span, a'));"
                "  var el = btns.find(function(e){"
                "    var t = (e.textContent || '').replace(/[\s　]/g,'');"
                "    return (t==='确定'||t==='確定') && e.offsetParent!==null;"
                "  });"
                "  if (el) { el.click(); return 'clicked:'+el.tagName+':'+el.textContent.trim(); }"
                "  return 'not found';"
                "}"
            )
            downloaded_file = None
            try:
                with lp_page.expect_download(timeout=30000) as dl_info:
                    confirmed = 'not found'
                    for sf in [lp_page, bi_frame] + list(lp_page.frames):
                        try:
                            r = sf.evaluate(JS_CONFIRM)
                            if str(r).startswith('clicked'):
                                confirmed = r
                                print(f"  Confirm: {r}")
                                break
                        except Exception:
                            pass
                    if not str(confirmed).startswith('clicked'):
                        print(f"  WARNING: confirm not found - check iur_after_export_menu.png")
                dl = dl_info.value
                save_path = os.path.join(DOWNLOAD_DIR, dl.suggested_filename or "iur_export.xlsx")
                dl.save_as(save_path)
                downloaded_file = save_path
                print(f"  Downloaded: {save_path}")
            except Exception as e:
                print(f"  Download wait failed: {e}")
                time.sleep(5)
                candidates = sorted(glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx")),
                                    key=os.path.getmtime, reverse=True)
                if candidates:
                    downloaded_file = candidates[0]
                    print(f"  Fallback file: {downloaded_file}")

            # Step 6: Save to rawdata
            if downloaded_file and os.path.exists(downloaded_file):
                print("Step 6: Saving to rawdata.xlsx sheet 2...")
                df = pd.read_excel(downloaded_file, header=0)
                print(f"  Shape: {df.shape[0]} rows x {df.shape[1]} cols")
                print(f"  Columns: {list(df.columns)}")
                print(df.head(3).to_string())
                with pd.ExcelWriter(rawdata_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df.to_excel(writer, sheet_name="2", index=False)
                print(f"  Saved to sheet 2 in {rawdata_file}")
            else:
                print("  ERROR: No downloaded file. Check screenshots in Output/")


        except Exception as e:
            print("ERROR: " + str(e))
            import traceback
            traceback.print_exc()
            try:
                tgt = lp_page if lp_page else page
                err_path = os.path.join(parent_dir, "Output",
                           "error_iur_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".png")
                tgt.screenshot(path=err_path, full_page=True)
                print("Screenshot: " + err_path)
            except Exception:
                pass
        finally:
            context.close()
            print("DONE")


if __name__ == "__main__":
    scrape_iur_new_report()
