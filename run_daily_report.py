# -*- coding: utf-8 -*-
"""
run_daily_report.py
Unified daily pipeline: scrape CRM + AMS -> generate CM and EA dashboards -> send Lark cards
"""
import os
import sys
import json
import shutil
import subprocess
import time
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

if sys.platform == "win32":
    try:
        if sys.stdout is not None:
            sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# -- Paths ------------------------------------------------------------------
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR    = os.path.join(SCRIPT_DIR, "Input")
OUTPUT_DIR   = os.path.join(SCRIPT_DIR, "Output")
RAWDATA_FILE = os.path.join(INPUT_DIR, "rawdata.xlsx")

# -- Read server config (credentials + active webhook) ----------------------
def _read_cfg():
    try:
        with open(os.path.join(SCRIPT_DIR, "data", "config.json")) as f:
            return json.load(f)
    except Exception:
        return {}

_cfg = _read_cfg()

TEST_MODE = "--test" in sys.argv

# -- Lark credentials -------------------------------------------------------
LARK_APP_ID       = "cli_a9bf7d0d8438dbdc"
LARK_APP_SECRET   = "fLNIH2ElbH9mChpijh4tbeKd36dJHKtq"
LARK_CHAT_ID_PROD = "oc_cc12fe7005d8a9fa8b8eb51e9193eeec"
LARK_CHAT_ID_TEST = "oc_1ab849cf11a8505ae909eff1928cd052"
LARK_CHAT_ID      = LARK_CHAT_ID_TEST if TEST_MODE else LARK_CHAT_ID_PROD

# -- DingTalk config --------------------------------------------------------
DINGTALK_WEBHOOK_URL = (
    _cfg.get("active_webhook_url") or
    "https://oapi.dingtalk.com/robot/send?access_token=28bc378d0fc40e94d1ae14f3223373c8d6fe6654e6595dd4ff6a138ecc3de0a3"
)
# Images served directly from local nginx — no GitHub needed
LOCAL_IMAGE_BASE      = "https://ahmed-live-lab-u56467.vm.elestio.app:15011/Output"

# -- Import local scripts ---------------------------------------------------
sys.path.insert(0, os.path.join(SCRIPT_DIR, "Scripts"))
sys.path.insert(0, SCRIPT_DIR)
from scrape_iur_report import scrape_iur_new_report
from scrape_crm_report import scrape_crm_report
from generate_cm_report import (
                                 read_duration_data, read_iur_data, read_cm_structure, merge_all_data,
                                 generate_html_individual_report, generate_html_separate_teams_report,
                                 generate_html_bottom20_report, generate_screenshots)
from html_report_generator import generate_html_team_report
from generate_ea_report import read_ea_structure, aggregate_monthly_data, merge_ea_data


# -- Phase 1: CRM scraping --------------------------------------------------
def run_crm_scrape():
    print("=" * 60)
    print("PHASE 1: Scraping CRM (crm.51talk.com)")
    print("=" * 60)
    try:
        scrape_crm_report()
        print("OK CRM scraping done.")
    except Exception as e:
        print(f"WARNING: CRM scraping failed: {e}")
        print("Continuing with existing rawdata.xlsx tab 1...")


def run_ams_scrape():
    print("=" * 60)
    print("PHASE 2: Scraping AMS (ams.51talkjr.com)")
    print("=" * 60)
    try:
        scrape_iur_new_report()
        print("OK AMS scraping done.")
    except Exception as e:
        print(f"WARNING: AMS scraping failed: {e}")
        print("Continuing with existing rawdata.xlsx tab 2...")


def run_cm_dashboard():
    print("=" * 60)
    print("PHASE 3: Generating CM Dashboard")
    print("=" * 60)
    structure = os.path.join(INPUT_DIR, "Team Structure.xlsx")
    dur_df  = read_duration_data(RAWDATA_FILE)
    iur_df  = read_iur_data(RAWDATA_FILE)
    str_df  = read_cm_structure(structure)
    merged  = merge_all_data(dur_df, iur_df, str_df)

    team_summary_html = os.path.join(OUTPUT_DIR, "CM_Team_Summary.html")
    sep_teams_html    = os.path.join(OUTPUT_DIR, "CM_Separate_Teams.html")

    generate_html_team_report(merged, team_summary_html)
    generate_html_separate_teams_report(merged, sep_teams_html)

    html_files = [team_summary_html, sep_teams_html]
    generate_screenshots(html_files, OUTPUT_DIR)

    pngs = [
        os.path.join(OUTPUT_DIR, "CM_Team_Summary.png"),
        os.path.join(OUTPUT_DIR, "CM_Separate_Teams.png"),
    ]
    print("OK CM dashboard done.")
    return merged, pngs


def run_ea_dashboard():
    print("=" * 60)
    print("PHASE 4: Generating EA Dashboard")
    print("=" * 60)
    structure     = os.path.join(INPUT_DIR, "Team Structure.xlsx")
    monthly_files = [os.path.join(INPUT_DIR, "EA_rawdata_Nov_Jan.xlsx")]

    ea_str  = read_ea_structure(structure)
    dur_df  = aggregate_monthly_data(monthly_files)
    merged  = merge_ea_data(dur_df, ea_str)

    individual_html   = os.path.join(OUTPUT_DIR, "EA_Individual_Report.html")
    team_summary_html = os.path.join(OUTPUT_DIR, "EA_Team_Summary.html")
    sep_teams_html    = os.path.join(OUTPUT_DIR, "EA_Separate_Teams.html")
    bottom20_html     = os.path.join(OUTPUT_DIR, "EA_Bottom20.html")

    generate_html_individual_report(merged, individual_html)
    generate_html_team_report(merged, team_summary_html)
    generate_html_separate_teams_report(merged, sep_teams_html)
    generate_html_bottom20_report(merged, bottom20_html)

    html_files = [individual_html, team_summary_html, sep_teams_html, bottom20_html]
    generate_screenshots(html_files, OUTPUT_DIR)

    pngs = [
        os.path.join(OUTPUT_DIR, "EA_Team_Summary.png"),
        os.path.join(OUTPUT_DIR, "EA_Separate_Teams.png"),
        os.path.join(OUTPUT_DIR, "EA_Individual_Report.png"),
        os.path.join(OUTPUT_DIR, "EA_Bottom20.png"),
    ]
    print("OK EA dashboard done.")
    return merged, pngs





# -- Phase 6: DingTalk webhook ----------------------------------------------
def get_local_image_urls(png_paths):
    """Return local nginx URLs for the given PNG paths (no GitHub needed)."""
    url_map = {}
    for path in png_paths:
        if os.path.exists(path):
            fname = os.path.basename(path)
            url_map[fname] = f"{LOCAL_IMAGE_BASE}/{fname}"
            print(f"  Image URL: {url_map[fname]}")
    return url_map


def dingtalk_send_webhook(image_urls, labels):
    """Send markdown card with images to DingTalk via webhook."""
    today = datetime.now().strftime("%Y-%m-%d")
    title = "CM Duration Report - {} {} [operation]".format(today, datetime.now().strftime("%H:%M"))

    lines = ["## " + title, ""]
    for label, url in zip(labels, image_urls):
        lines += ["**" + label + "**", "", "![](" + url + ")", "", "---", ""]
    text = chr(10).join(lines)

    payload = {
        "msgtype": "markdown",
        "markdown": {"title": title, "text": text}
    }
    resp = requests.post(DINGTALK_WEBHOOK_URL, json=payload, timeout=15)
    result = resp.json()
    if result.get("errcode") != 0:
        raise RuntimeError("DingTalk send failed: {}".format(result))
    print("  DingTalk message sent.")

def run_send_dingtalk(cm_pngs):
    print("=" * 60)
    print("PHASE 6: Sending to DingTalk")
    print("=" * 60)
    try:
        labels = ["Team Summary", "Teams Breakdown"]
        valid_pairs = [(l, p) for l, p in zip(labels, cm_pngs) if os.path.exists(p)]
        valid_labels = [l for l, _ in valid_pairs]
        valid_paths  = [p for _, p in valid_pairs]

        print("  Building local image URLs...")
        url_map = get_local_image_urls(valid_paths)
        image_urls = [url_map[os.path.basename(p)] for p in valid_paths]

        print("  Sending DingTalk webhook...")
        dingtalk_send_webhook(image_urls, valid_labels)
        print("OK DingTalk sent.")
    except Exception as e:
        print(f"WARNING: DingTalk send failed: {e}")


# -- Phase 5: Send Lark cards -----------------------------------------------
def lark_get_token():
    resp = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": LARK_APP_ID, "app_secret": LARK_APP_SECRET},
        timeout=10
    )
    result = resp.json()
    if result.get("code") == 0:
        return result["tenant_access_token"]
    raise RuntimeError(f"Lark auth failed: {result}")


def lark_upload_image(token, image_path):
    headers = {"Authorization": f"Bearer {token}"}
    with open(image_path, "rb") as f:
        resp = requests.post(
            "https://open.feishu.cn/open-apis/im/v1/images",
            headers=headers,
            files={"image": (os.path.basename(image_path), f, "image/png")},
            data={"image_type": "message"},
            timeout=30
        )
    result = resp.json()
    if result.get("code") == 0:
        return result["data"]["image_key"]
    raise RuntimeError(f"Upload failed for {image_path}: {result}")


def lark_send_card(token, chat_id, title, color, image_keys, labels):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    elements = []
    for label, key in zip(labels, image_keys):
        if key:
            elements.append({"tag": "div", "text": {"tag": "lark_md", "content": f"**{label}**"}})
            elements.append({"tag": "img", "img_key": key, "alt": {"tag": "plain_text", "content": label}})
            elements.append({"tag": "hr"})

    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "title": {"tag": "plain_text", "content": title},
            "template": color
        },
        "elements": elements
    }
    payload = {
        "receive_id": chat_id,
        "msg_type": "interactive",
        "content": json.dumps(card)
    }
    resp = requests.post(
        "https://open.feishu.cn/open-apis/im/v1/messages",
        headers=headers,
        params={"receive_id_type": "chat_id"},
        json=payload,
        timeout=10
    )
    result = resp.json()
    if result.get("code") != 0:
        raise RuntimeError(f"Send failed: {result}")
    msg_id = result.get("data", {}).get("message_id", "unknown")
    print(f"  Message ID: {msg_id}")
    return msg_id


def run_send_cards(cm_pngs, ea_pngs):
    print("=" * 60)
    print("PHASE 5: Sending cards to Lark group")
    print("=" * 60)
    for attempt in range(1, 4):
        try:
            token = lark_get_token()
            print(f"  Token obtained (attempt {attempt})")

            today = datetime.now().strftime("%Y-%m-%d")

            # CM card
            cm_labels = ["Team Summary", "Teams Breakdown"]
            print("  Uploading CM images...")
            valid_pairs = [(l, p) for l, p in zip(cm_labels, cm_pngs) if os.path.exists(p)]
            cm_valid_labels = [l for l, _ in valid_pairs]
            valid_paths = [p for _, p in valid_pairs]
            with ThreadPoolExecutor(max_workers=len(valid_paths) or 1) as pool:
                cm_keys = list(pool.map(lambda p: lark_upload_image(token, p), valid_paths))
            lark_send_card(token, LARK_CHAT_ID,
                           f"CM Duration Report - {today} {datetime.now().strftime('%H:%M')}", "blue",
                           cm_keys, cm_valid_labels)
            print("  CM card sent")

            # EA card (commented out)
            # ea_labels = ["Team Summary", "Teams Breakdown", "Individual Report", "Bottom 20"]
            # print("  Uploading EA images...")
            # ea_keys = [lark_upload_image(token, p) for p in ea_pngs if os.path.exists(p)]
            # ea_valid_labels = [l for l, p in zip(ea_labels, ea_pngs) if os.path.exists(p)]
            # lark_send_card(token, LARK_CHAT_ID,
            #                f"EA Daily Report - {today}", "green",
            #                ea_keys, ea_valid_labels)
            # print("  EA card sent")

            print("OK Cards sent.")
            return
        except Exception as e:
            print(f"WARNING: Send attempt {attempt}/3 failed: {e}")
            if attempt < 3:
                time.sleep(10)
    print("ERROR: All 3 send attempts failed.")

if __name__ == "__main__":
    # --send-only: skip scraping/generation, just send existing PNGs
    if "--send-only" in sys.argv:
        cm_pngs = [
            os.path.join(OUTPUT_DIR, "CM_Team_Summary.png"),
            os.path.join(OUTPUT_DIR, "CM_Separate_Teams.png"),
        ]
        run_send_dingtalk(cm_pngs)
        sys.exit(0)

    if datetime.now().weekday() == 4:  # 4 = Friday
        print("Friday - no report today.")
        sys.exit(0)

    # -- Kill any previous instance (avoids Chrome profile lock conflicts) ----
    import psutil
    current_pid = os.getpid()
    for proc in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            if proc.pid == current_pid:
                continue
            if "python" not in proc.name().lower():
                continue
            cmdline = " ".join(proc.cmdline() or [])
            if "run_daily_report" in cmdline:
                for child in proc.children(recursive=True):
                    try: child.kill()
                    except Exception: pass
                proc.kill()
                print(f"Killed previous run (PID {proc.pid})")
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    log_dir = os.path.join(OUTPUT_DIR, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, datetime.now().strftime("%Y-%m-%d_%H-%M") + ".log")
    log_file = open(log_path, "w", encoding="utf-8")
    log_file.write("Log: " + log_path + "\n")
    log_file.flush()
    class _Tee:
        def __init__(self, *streams): self.streams = streams
        def write(self, data):
            for s in self.streams:
                try: s.write(data)
                except Exception: pass
        def flush(self):
            for s in self.streams:
                try: s.flush()
                except Exception: pass
    sys.stdout = _Tee(sys.stdout, log_file)
    print(f"Log: {log_path}")
    run_crm_scrape()
    run_ams_scrape()
    cm_df, cm_pngs = run_cm_dashboard()
    # run_send_cards(cm_pngs, [])  # Lark: commented out
    run_send_dingtalk(cm_pngs)
    sys.stdout = sys.stdout.streams[0]
    log_file.close()

