# -*- coding: utf-8 -*-
"""
setup_schedule.py
Run ONCE as Administrator to register all scrap_bot Task Scheduler tasks.
"""
import subprocess, os

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
BAT_SCRIPT  = os.path.join(SCRIPT_DIR, "run_daily_report.bat")
FOLDER      = "scrap_bot"
WEEKDAYS    = "SUN,MON,TUE,WED,THU"  # skip Friday and Saturday

TASKS = [
    ("scrap_bot_1500", "15:00"),
    ("scrap_bot_1800", "18:00"),
    ("scrap_bot_2100", "21:00"),
]

def create_task(name, time_str):
    args = [
        "schtasks", "/Create", "/F",
        "/TN", FOLDER + chr(92) + name,
        "/TR", 'cmd.exe /c "' + BAT_SCRIPT + '"',
        "/SC", "WEEKLY",
        "/D", WEEKDAYS,
        "/ST", time_str,
    ]
    r = subprocess.run(args, capture_output=True, text=True)
    status = "OK" if r.returncode == 0 else r.stderr.strip()
    print(f"  {name} @ {time_str}: {status}")

if __name__ == "__main__":
    print("=" * 50)
    print("scrap_bot Scheduler Setup")
    print("=" * 50)
    print(f"Batch:   {BAT_SCRIPT}")
    print(f"Days:    {WEEKDAYS}")
    print()
    for name, t in TASKS:
        create_task(name, t)
    print()
    print("Done. Verify in Task Scheduler (taskschd.msc) under:", FOLDER)
    print("To remove all tasks: python remove_schedule.py")
