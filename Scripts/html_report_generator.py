import pandas as pd
import os
import sys
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
def generate_html_team_report(merged_df, output_file):
    """Generate HTML team summary report"""

    team_summary = merged_df.groupby("Team").agg({
        "Name": "count",
        "Total Eff. Calls": "sum",
        "Total Duration (Min)": "sum",
        "Classes Completed": "sum",
    }).reset_index()
    team_summary.rename(columns={"Name": "Members"}, inplace=True)

    team_summary["Avg Call Time/Min"] = (team_summary["Total Duration (Min)"] / team_summary["Members"]).apply(lambda x: int(round(x)))
    team_summary["Avg Eff. Calls"] = (team_summary["Total Eff. Calls"] / team_summary["Members"]).apply(lambda x: int(round(x)))
    team_summary = team_summary[["Team", "Members", "Total Eff. Calls", "Total Duration (Min)", "Avg Eff. Calls", "Avg Call Time/Min", "Classes Completed"]]
    team_summary = team_summary.sort_values("Avg Call Time/Min", ascending=False)

    total_eff_calls = team_summary["Total Eff. Calls"].sum()
    total_duration  = team_summary["Total Duration (Min)"].sum()
    total_members   = team_summary["Members"].sum()

    total_row = pd.DataFrame([{
        "Team": "TOTAL",
        "Members": total_members,
        "Total Eff. Calls": total_eff_calls,
        "Total Duration (Min)": total_duration,
        "Avg Eff. Calls": int(round(team_summary["Avg Eff. Calls"].mean())) if total_members > 0 else 0,
        "Avg Call Time/Min": int(round(team_summary["Avg Call Time/Min"].mean())) if total_members > 0 else 0,
        "Classes Completed": team_summary["Classes Completed"].sum(),
    }])
    team_summary = pd.concat([team_summary, total_row], ignore_index=True)

    data_rows = team_summary[team_summary["Team"] != "TOTAL"]
    non_zero_rows = data_rows[data_rows["Total Duration (Min)"] > 0]
    if len(non_zero_rows) > 0:
        duration_min = non_zero_rows["Total Duration (Min)"].min()
        duration_max = non_zero_rows["Total Duration (Min)"].max()
        avg_call_min = non_zero_rows["Avg Call Time/Min"].min()
        avg_call_max = non_zero_rows["Avg Call Time/Min"].max()
    else:
        duration_min = data_rows["Total Duration (Min)"].min()
        duration_max = data_rows["Total Duration (Min)"].max()
        avg_call_min = data_rows["Avg Call Time/Min"].min()
        avg_call_max = data_rows["Avg Call Time/Min"].max()

    def get_color_class(value, min_val, max_val):
        if value == 0:
            return "scale-very-low"
        if max_val == 0:
            return "scale-very-low"
        normalized = value / max_val
        if normalized < 0.17:   return "scale-very-low"
        elif normalized < 0.33: return "scale-low"
        elif normalized < 0.50: return "scale-medium-low"
        elif normalized < 0.67: return "scale-medium"
        elif normalized < 0.83: return "scale-medium-high"
        else:                   return "scale-high"

    css = "* { margin: 0; padding: 0; box-sizing: border-box; }body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f7fa; padding: 30px; }.container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 12px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; }h1 { font-size: 24px; font-weight: 600; color: #1a1a1a; }.download-btn { background: #f5f7fa; border: 1px solid #e1e4e8; padding: 8px 16px; border-radius: 6px; color: #586069; font-size: 14px; cursor: pointer; display: flex; align-items: center; gap: 8px; text-decoration: none; }.download-btn:hover { background: #e9ecef; }table { width: 100%; border-collapse: separate; border-spacing: 0; }thead th { background: #f4f6f8; color: #64748b; font-weight: 500; font-size: 15px; text-align: left; padding: 12px 16px; border-bottom: 1px solid #e2e8f0; }tbody td { padding: 12px 16px; border-bottom: 1px solid #e2e8f0; font-size: 16px; color: #1a365d; }tbody tr:hover { background: #f4f6f8; }.total-row td { font-weight: 600; background: #f4f6f8; color: #1a365d; }.badge { display: inline-block; padding: 6px 14px; border-radius: 20px; font-weight: 500; font-size: 13px; text-align: center; min-width: 60px; }.scale-high { background: #63BE7B; color: #000; }.scale-medium-high { background: #9FD899; color: #000; }.scale-medium { background: #C6E5B5; color: #000; }.scale-medium-low { background: #FFEB84; color: #000; }.scale-low { background: #FCAA75; color: #000; }.scale-very-low { background: #F8696B; color: #000; }.text-center { text-align: center !important; }"

    html = (
        '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
        '<title>Report by Teams Totals</title><style>' + css + '</style></head>'
        '<body><div class="container" id="teamsTable">'
        '<div class="header"><h1>Report by Teams Totals</h1>'
        '<a class="download-btn" href="#">&#8595; Download PNG</a></div>'
        '<table><thead><tr>'
        '<th>Team</th>'
        '<th class="text-center">Members</th>'
        '<th class="text-center">Total Eff. Calls</th>'
        '<th class="text-center">Total Duration (Min)</th>'
        '<th class="text-center">Avg Eff. Calls</th>'
        '<th class="text-center">Avg Call Time/Min</th>'
        '<th class="text-center">Classes Completed</th>'
        '</tr></thead><tbody>'
    )

    for idx, row in team_summary.iterrows():
        team_name = row["Team"]
        is_total  = team_name == "TOTAL"

        duration_color = "" if is_total else get_color_class(row["Total Duration (Min)"], duration_min, duration_max)
        avg_call_color = "" if is_total else get_color_class(row["Avg Call Time/Min"], avg_call_min, avg_call_max)

        eff_calls_val = int(row["Total Eff. Calls"])
        duration_val  = int(row["Total Duration (Min)"])
        eff_calls_display = f"{eff_calls_val:,}" if is_total else str(eff_calls_val)
        duration_display  = f"{duration_val:,}" if is_total else str(duration_val)

        row_class = ' class="total-row"' if is_total else ""
        html += (
            f'<tr{row_class}>'
            f'<td><strong>{team_name}</strong></td>'
            f'<td class="text-center">{row["Members"]}</td>'
            f'<td class="text-center">{eff_calls_display}</td>'
            f'<td class="text-center"><span class="badge {duration_color}">{duration_display}</span></td>'
            f'<td class="text-center">{row["Avg Eff. Calls"]}</td>'
            f'<td class="text-center"><span class="badge {avg_call_color}">{row["Avg Call Time/Min"]}</span></td>'
            f'<td class="text-center">{int(row["Classes Completed"])}</td>'
            f'</tr>'
        )

    html += '</tbody></table></div></body></html>'

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"HTML report saved to: {output_file}")

if __name__ == "__main__":
    pass
