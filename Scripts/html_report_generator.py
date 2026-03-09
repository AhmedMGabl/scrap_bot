import pandas as pd
import os

def generate_html_team_report(merged_df, output_file):
    """Generate HTML team summary report matching the screenshot style"""
    
    # Group by team and calculate aggregates
    team_summary = merged_df.groupby('Team').agg({
        'Name': 'count',
        'Total Eff. Calls': 'sum',
        'Total Duration (Min)': 'sum',
        'Classes Completed': 'sum',
    }).reset_index()

    # Avg Call Time/Min per team = sum of each agent's rounded individual avg call time
    avg_call_by_team = merged_df.groupby('Team')['Avg Call Time/Min'].apply(
        lambda vals: sum(int(round(v)) for v in vals)
    ).rename('Avg Call Time/Min')
    team_summary = team_summary.merge(avg_call_by_team, on='Team')

    # Calculate averages
    team_summary['Avg Eff. Calls'] = (team_summary['Total Eff. Calls'] / team_summary['Name']).round(1)
    
    # Rename columns
    team_summary.columns = ['Team', 'Members', 'Total Eff. Calls', 'Total Duration (Min)',
                            'Classes Completed', 'Avg Eff. Calls', 'Avg Call Time/Min']
    
    # Reorder columns
    team_summary = team_summary[['Team', 'Members', 'Total Eff. Calls', 'Total Duration (Min)', 
                                 'Avg Eff. Calls', 'Avg Call Time/Min', 'Classes Completed']]
    
    # Sort by Avg Call Time/Min descending (best performers first)
    team_summary = team_summary.sort_values('Avg Call Time/Min', ascending=False)
    
    # Add total row
    total_eff_calls = team_summary['Total Eff. Calls'].sum()
    total_duration = team_summary['Total Duration (Min)'].sum()
    total_members = team_summary['Members'].sum()
    
    total_row = pd.DataFrame([{
        'Team': 'TOTAL',
        'Members': total_members,
        'Total Eff. Calls': total_eff_calls,
        'Total Duration (Min)': total_duration,
        'Avg Eff. Calls': round(total_eff_calls / total_members, 1) if total_members > 0 else 0,
        'Avg Call Time/Min': int(team_summary[team_summary['Team'] != 'TOTAL']['Avg Call Time/Min'].sum()),
        'Classes Completed': team_summary['Classes Completed'].sum(),
    }])
    
    team_summary = pd.concat([team_summary, total_row], ignore_index=True)
    
    # Calculate min/max for dynamic color scaling (exclude TOTAL row)
    data_rows = team_summary[team_summary['Team'] != 'TOTAL']
    duration_min = data_rows['Total Duration (Min)'].min()
    duration_max = data_rows['Total Duration (Min)'].max()
    avg_call_min = data_rows['Avg Call Time/Min'].min()
    avg_call_max = data_rows['Avg Call Time/Min'].max()

    # Function to get color class based on value position in range (6-level gradient)
    def get_color_class(value, min_val, max_val):
        """Excel-style color scaling: min=red, mid=yellow, max=green"""
        if max_val == min_val:
            return 'bg-medium'
        
        # Normalize value to 0-1 range
        normalized = (value - min_val) / (max_val - min_val)
        
        # Excel 3-color scale: 
        # 0.0 = red (bottom)
        # 0.5 = yellow (middle) 
        # 1.0 = green (top)
        if normalized < 0.5:
            # Lower half: red to yellow gradient
            # Map 0-0.5 to our 3 red/orange/yellow levels
            if normalized < 0.17:
                return 'bg-very-low'  # Pure red
            elif normalized < 0.33:
                return 'bg-low'  # Orange
            else:
                return 'bg-medium-low'  # Yellow
        else:
            # Upper half: yellow to green gradient
            # Map 0.5-1.0 to our 3 yellow/light green/dark green levels
            if normalized < 0.67:
                return 'bg-medium'  # Light green
            elif normalized < 0.83:
                return 'bg-medium-high'  # Medium green
            else:
                return 'bg-high'  # Dark green
    
    # Generate HTML
    html = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Report by Teams Totals</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: #f5f7fa;
            padding: 30px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            padding: 30px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
        }
        
        h1 {
            font-size: 24px;
            font-weight: 600;
            color: #1a1a1a;
        }
        
        .download-btn {
            background: #f5f7fa;
            border: 1px solid #e1e4e8;
            padding: 8px 16px;
            border-radius: 6px;
            color: #586069;
            font-size: 14px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .download-btn:hover {
            background: #e9ecef;
        }
        
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
        }
        
        thead th {
            background: #f5f7fa;
            color: #6b7280;
            font-weight: 600;
            font-size: 13px;
            text-align: left;
            padding: 14px 16px;
            border-bottom: 2px solid #e5e7eb;
        }
        
        tbody td {
            padding: 16px;
            border-bottom: 1px solid #f3f4f6;
            font-size: 14px;
            color: #1f2937;
        }
        
        tbody tr:hover {
            background: #fafbfc;
        }
        
        tbody tr:last-child {
            font-weight: 700;
            background: #dbeafe;
            color: #1e3a5f;
        }
        
        tbody tr:last-child td {
            border-bottom: none;
            border-top: 2px solid #3b82f6;
        }
        
        .badge {
            display: inline-block;
            padding: 6px 14px;
            border-radius: 20px;
            font-weight: 500;
            font-size: 13px;
            text-align: center;
            min-width: 60px;
        }
        
        .bg-high {
            background: #63BE7B;
            color: #000;
        }
        
        .bg-medium-high {
            background: #9FD899;
            color: #000;
        }
        
        .bg-medium {
            background: #C6E5B5;
            color: #000;
        }
        
        .bg-medium-low {
            background: #FFEB84;
            color: #000;
        }
        
        .bg-low {
            background: #FCAA75;
            color: #000;
        }
        
        .bg-very-low {
            background: #F8696B;
            color: #000;
        }
        
        .text-center {
            text-align: center !important;
        }
        
        @media print {
            body {
                padding: 0;
                background: white;
            }
            
            .container {
                box-shadow: none;
                padding: 20px;
            }
            
            .download-btn {
                display: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Report by Teams Totals</h1>
        </div>
        
        <table>
            <thead>
                <tr>
                    <th>Team</th>
                    <th class="text-center">Members</th>
                    <th class="text-center">Total Eff. Calls</th>
                    <th class="text-center">Total Duration (Min)</th>
                    <th class="text-center">Avg Eff. Calls</th>
                    <th class="text-center">Avg Call Time/Min</th>
                    <th class="text-center">Classes Completed</th>
                    
                </tr>
            </thead>
            <tbody>
"""
    
    # Add data rows
    for idx, row in team_summary.iterrows():
        team_name = row['Team']
        is_total = team_name == 'TOTAL'
        
        duration_color = '' if is_total else get_color_class(row['Total Duration (Min)'], duration_min, duration_max)
        avg_call_color = '' if is_total else get_color_class(row['Avg Call Time/Min'], avg_call_min, avg_call_max)
        
        html += f"""
                <tr>
                    <td><strong>{team_name}</strong></td>
                    <td class="text-center">{row['Members']}</td>
                    <td class="text-center">{row['Total Eff. Calls']}</td>
                    <td class="text-center"><span class="badge {duration_color}">{row['Total Duration (Min)']}</span></td>
                    <td class="text-center">{row['Avg Eff. Calls']}</td>
                    <td class="text-center"><span class="badge {avg_call_color}">{row['Avg Call Time/Min']}</span></td>
                    <td class="text-center">{row['Classes Completed']}</td>
                </tr>
"""
    
    html += """
            </tbody>
        </table>
    </div>
</body>
</html>
"""
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"HTML report saved to: {output_file}")

if __name__ == "__main__":
    # This can be called from the main script
    pass
