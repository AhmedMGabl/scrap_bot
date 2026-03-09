# -*- coding: utf-8 -*-
"""remove_schedule.py - Removes all scrap_bot Task Scheduler tasks."""
import subprocess

FOLDER = "scrap_bot"
TASKS = [
    "scrap_bot_morning_1230", "scrap_bot_morning_1330",
    "scrap_bot_morning_1430", "scrap_bot_morning_1530",
    "scrap_bot_night_2200",   "scrap_bot_night_2300",
    "scrap_bot_night_0000",   "scrap_bot_saturday_late",
]

if __name__ == "__main__":
    print("Removing scrap_bot scheduled tasks...")
    for name in TASKS:
        r = subprocess.run(
            f"schtasks /Delete /F /TN "{FOLDER}\{name}"",
            capture_output=True, text=True, shell=True
        )
        print(f"  {name}: {'OK' if r.returncode == 0 else r.stderr.strip()}")
    print("Done.")
