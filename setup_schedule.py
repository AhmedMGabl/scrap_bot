# -*- coding: utf-8 -*-
"""
setup_schedule.py
Run ONCE as Administrator to register all scrap_bot Task Scheduler tasks.
"""
import subprocess, os, sys

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
PYTHON_EXE  = sys.executable
MAIN_SCRIPT = os.path.join(SCRIPT_DIR, "run_daily_report.py")
FOLDER      = "scrap_bot"

TASKS = [
    ("scrap_bot_morning_1230", "12:30", "DAILY",  None),
    ("scrap_bot_morning_1330", "13:30", "DAILY",  None),
    ("scrap_bot_morning_1430", "14:30", "DAILY",  None),
    ("scrap_bot_morning_1530", "15:30", "DAILY",  None),
    ("scrap_bot_night_2200",   "22:00", "DAILY",  None),
    ("scrap_bot_night_2300",   "23:00", "DAILY",  None),
    ("scrap_bot_night_0000",   "00:00", "DAILY",  None),
    ("scrap_bot_saturday_late","01:00", "WEEKLY", "SUN"),
]

def create_task(name, time_str, schedule, day):
    # /TR must wrap both paths in escaped inner quotes to handle spaces
    tr_inner = chr(92) + chr(34) + PYTHON_EXE + chr(92) + chr(34) + " " + chr(92) + chr(34) + MAIN_SCRIPT + chr(92) + chr(34)
    tr = chr(34) + tr_inner + chr(34)
    tn = chr(34) + FOLDER + chr(92) + name + chr(34)
    parts = ["schtasks", "/Create", "/F", "/TN", tn, "/TR", tr, "/SC", schedule]
    if day:
        parts += ["/D", day]
    parts += ["/ST", time_str]
    cmd = " ".join(parts)
    r = subprocess.run(cmd, capture_output=True, text=True, shell=True)
    status = "OK" if r.returncode == 0 else r.stderr.strip()
    print(f"  {name} @ {time_str}: {status}")

if __name__ == "__main__":
    print("=" * 50)
    print("scrap_bot Scheduler Setup")
    print("=" * 50)
    print(f"Python:  {PYTHON_EXE}")
    print(f"Script:  {MAIN_SCRIPT}")
    print()
    for args in TASKS:
        create_task(*args)
    print()
    print("Done. Verify in Task Scheduler (taskschd.msc) under:", FOLDER)
    print("To remove all tasks: python remove_schedule.py")
