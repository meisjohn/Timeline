import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os, sys, subprocess

# 1. Setup & Args
input_arg = sys.argv if len(sys.argv) > 1 else "data/events.xlsx"
output_dir = "outputs"
os.makedirs(output_dir, exist_ok=True)

STATUS_MAP = {
    "Late": {"border": "#c0392b", "fill": "#e74c3c", "width": 4},
    "Delayed": {"border": "#e67e22", "fill": None, "width": 4},
    "At Risk": {"border": "#f1c40f", "fill": None, "width": 3},
    "Complete": {"border": "#27ae60", "fill": None, "width": 2},
    "On Track": {"border": "white", "fill": None, "width": 1}
}

NON_RED_PALETTE = ["#1f77b4", "#2ca02c", "#ff7f0e", "#9467bd", "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]

# 2. Load and Prepare Data
try:
    df = pd.read_excel(input_arg)
    df['Start'], df['End'] = pd.to_datetime(df['Start']), pd.to_datetime(df['End'])
    df['Status'] = df['Status'].fillna('On Track').astype(str).str.strip()
    df = df[df['Include'].astype(str).str.strip().str.upper() == 'YES'].copy()
except Exception as e:
    print(f"Excel Error: {e}"); sys.exit(1)

df = df.sort_values(by='End', ascending=True)
summary_df = df.copy()
summary_df['Event Name'] = "ALL EVENTS" 

plot_df = pd.concat([summary_df, df], ignore_index=True)
event_order = df['Event Name'].unique().tolist()[::-1] + ["ALL EVENTS"]

# 3. Visualization Constants
today = datetime.now()
timestamp = today.strftime("%Y%m%d_%H%M%S")

# STRICT bounds to prevent the axis from stretching into 2027
data_min = plot_df['Start'].min() - timedelta(days=5)
data_max = plot_df['End'].max() + timedelta(days=5)

views = [
    {"name": "Full_Timeline", "range": [data_min, data_max]}, # Forces tight bounds
    {"name": "Rolling_Window", "range": [today - timedelta(days=30), today + timedelta(days=90)]}
]

for view in views:
    fig = px.timeline(
        plot_df, x_start="Start", x_end="End", y="Event Name", 
        color="Task/Phase",
        color_discrete_sequence=NON_RED_PALETTE,
        category_orders={"Event Name": event_order}
    )

    # 4. Apply Status Borders
    for trace in fig.data:
        trace_tasks = plot_df[plot_df['Task/Phase'] == trace.name]
        b_colors, b_widths = [], []
        f_colors = list(trace.marker.color) if isinstance(trace.marker.color, (list, tuple)) else [trace.marker.color] * len(trace_tasks)

        for i, status in enumerate(trace_tasks['Status']):
            s_cfg = STATUS_MAP.get(status, STATUS_MAP["On Track"])
            b_colors.append(s_cfg['border'])
            b_widths.append(s_cfg['width'])
            if s_cfg['fill']: f_colors[i] = s_cfg['fill']

        trace.marker.line.color = b_colors
        trace.marker.line.width = b_widths
        trace.marker.color = f_colors

    # 5. Legend
    for status, style in STATUS_MAP.items():
        fig.add_trace(go.Bar(x=[None], y=[None], name=f"Status: {status}",
                             marker=dict(color=style['fill'] if style['fill'] else 'rgba(0,0,0,0)', 
                                         line=dict(color=style['border'], width=style['width'])),
                             showlegend=True))

    # 6. Banded Columns (Vertical Stripes)
    stripe_start = data_min - timedelta(days=30)
    for i in range(0, 40): 
        fig.add_vrect(
            x0=stripe_start + timedelta(days=i*14), 
            x1=stripe_start + timedelta(days=(i*14)+7),
            fillcolor="#fbfbfb", layer="below", line_width=0
        )

    # 7. Overlays & Separators
    # Transparent box from far-left (relative to view) to Today
    box_left = view["range"][0] if view["range"] else data_min
    fig.add_vrect(x0=box_left - timedelta(days=200), x1=today, 
                  fillcolor="white", opacity=0.85, layer="above", line_width=0)
    
    # Thick separator below "ALL EVENTS"
    fig.add_shape(type="line", x0=0, x1=1, xref="paper", y0=0.5, y1=0.5, 
                  line=dict(color="black", width=4), layer="above")
    
    # Today Marker
    fig.add_shape(type="line", x0=today, x1=today, y0=0, y1=1, yref="paper",
                  line=dict(color="#d63031", width=2), layer="above")
    
    fig.add_annotation(x=today, y=0.05, yref="paper", 
                       text=f" TODAY ({today.strftime('%b %d')})", 
                       showarrow=False, font=dict(color="#d63031", size=11, family="Arial Black"),
                       bgcolor="white", xanchor="left")

    # 8. Layout & Banded Rows
    fig.update_layout(
        height=max(400, (len(event_order) * 45) + 200),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=180, r=40, t=100, b=120),
        xaxis=dict(
            showgrid=True, gridcolor="#eeeeee", 
            tickformat="%b %d %Y", # Added Year
            side="top",
            mirror="allticks", ticks="outside",
            range=view["range"] # Forces strict axis clipping
        ),
        yaxis=dict(
            title="", autorange="reversed",
            showgrid=True, gridcolor="#eeeeee",
            zeroline=False
        ),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5)
    )

    filename = os.path.join(output_dir, f"{view['name']}_{timestamp}.png")
    fig.write_image(filename, width=1600, scale=3)

# 9. Open Folder
print(f"Success! Charts saved to {output_dir}")
def open_folder(path):
    if sys.platform == 'win32': os.startfile(os.path.abspath(path))
    elif sys.platform == 'darwin': subprocess.Popen(['open', path])
    else: subprocess.Popen(['xdg-open', path])
open_folder(output_dir)
