import pandas as pd
from bokeh.plotting import figure
from bokeh.io import save
from bokeh.models import ColumnDataSource, Span, Label, BoxAnnotation, Range1d, DatetimeTickFormatter, Legend, LegendItem, HoverTool
from datetime import datetime, timedelta
import os, sys, subprocess
from io import BytesIO

# 1. Setup & Args
input_arg = "data/events.xlsx"
for arg in sys.argv[1:]:
        input_arg = arg
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
summary_df['Y_Axis_Label'] = "ALL EVENTS"
df['Y_Axis_Label'] = df['Event Name']

plot_df = pd.concat([summary_df, df], ignore_index=True)
event_order = ["ALL EVENTS"] + df['Event Name'].unique().tolist()[::-1]

# Pre-calculate Bokeh Styles (Colors & Borders)
unique_phases = plot_df['Task/Phase'].unique()
phase_map = {p: NON_RED_PALETTE[i % len(NON_RED_PALETTE)] for i, p in enumerate(unique_phases)}
plot_df['base_color'] = plot_df['Task/Phase'].map(phase_map)

def get_style(row):
    s = STATUS_MAP.get(row['Status'], STATUS_MAP["On Track"])
    # If status has no fill, use the Task/Phase color
    fc = s['fill'] if s['fill'] else row['base_color']
    return pd.Series([s['border'], fc, s['width']])
plot_df[['b_line', 'b_fill', 'b_width']] = plot_df.apply(get_style, axis=1)

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
    # Filter Data for Rolling Window
    if view["name"] == "Rolling_Window":
        w_start, w_end = view["range"]
        # Check for overlap: Task Start < Window End AND Task End > Window Start
        mask = (plot_df['Start'] < w_end) & (plot_df['End'] > w_start)
        valid_events = plot_df.loc[mask, 'Event Name'].unique()
        view_df = plot_df[plot_df['Event Name'].isin(valid_events)].copy()
        view_event_order = [e for e in event_order if e in valid_events or e == "ALL EVENTS"]
    else:
        view_df = plot_df
        view_event_order = event_order

    # Initialize Bokeh Figure
    # Note: We reverse event_order for y_range to match Plotly's top-to-bottom layout
    p = figure(
        y_range=view_event_order[::-1], 
        x_range=Range1d(view["range"][0], view["range"][1]),
        x_axis_type="datetime",
        height=max(400, (len(view_event_order) * 45) + 200),
        width=1600,
        title=view["name"].replace("_", " "),
        toolbar_location="above",
        tools="pan,wheel_zoom,box_zoom,reset,save"
    )
    p.title.align = "center"
    p.title.text_font_size = "16pt"

    # 4. Banded Columns (Vertical Stripes)
    stripe_start = data_min - timedelta(days=30)
    for i in range(0, 40): 
        s = stripe_start + timedelta(days=i*14)
        p.add_layout(BoxAnnotation(left=s, right=s+timedelta(days=7), fill_color="#e0e0e0", fill_alpha=1, level="image"))

    # 5. Overlays & Separators
    # Transparent box from far-left (relative to view) to Today
    box_left = view["range"][0] if view["range"] else data_min
    p.add_layout(BoxAnnotation(left=box_left - timedelta(days=200), right=today, fill_color="white", fill_alpha=0.75, level="overlay"))
    
    # Today Marker
    p.add_layout(Span(location=today, dimension='height', line_color="#d63031", line_width=2, level="overlay"))
    p.add_layout(Label(x=today, y=10, y_units='screen', text=f" TODAY ({today.strftime('%b %d')})", 
                       text_color="#d63031", text_font_style="bold", background_fill_color="white"))

    # Highlight "ALL EVENTS" Row Background
    if "ALL EVENTS" in view_event_order:
        p.hbar(y=["ALL EVENTS"], left=view["range"][0]-(view["range"][1]-view["range"][0]), right=view["range"][1]+(view["range"][1]-view["range"][0]), height=1, 
               fill_color=None, line_color="#000000", line_width=2, level="underlay")

    # 6. Main Timeline Bars
    src = ColumnDataSource(view_df)
    timeline_bars = p.hbar(y='Y_Axis_Label', left='Start', right='End', height=0.6, source=src,
                           fill_color='b_fill', line_color='b_line', line_width='b_width')

    # 7. Legend (Status)
    items = []
    for status, style in STATUS_MAP.items():
        # Create invisible dummy glyphs for the legend
        dummy = p.hbar(y=[view_event_order[0]] if view_event_order else [], left=[today], right=[today], visible=False,
                       fill_color=style['fill'] or "white", line_color=style['border'], line_width=style['width'])
        items.append(LegendItem(label=f"Status: {status}", renderers=[dummy]))
    p.add_layout(Legend(items=items, location="bottom_center", orientation="horizontal"), 'below')

    # 8. Styling
    p.xaxis.formatter = DatetimeTickFormatter(days="%b %d %Y", months="%b %d %Y")
    p.xaxis.axis_label = ""
    p.yaxis.axis_label = ""
    p.ygrid.grid_line_color = None
    p.outline_line_color = None

    # Add Tooltips
    p.add_tools(HoverTool(
        renderers=[timeline_bars],
        tooltips=[
            ("Event", "@{Event Name}"),
            ("Start", "@Start{%F}"),
            ("End", "@End{%F}"),
            ("Phase", "@{Task/Phase}"),
            ("Status", "@Status")
        ],
        formatters={'@Start': 'datetime', '@End': 'datetime'}
    ))

    # Save Image & Optional PowerPoint
    output_filename = os.path.join(output_dir, f"{view['name']}_{timestamp}.html")
    save(p, filename=output_filename)

output_path = output_dir

# 9. Open File
print(f"Success! Charts saved to {output_path}")
def open_folder(path):
    if sys.platform == 'win32': os.startfile(os.path.abspath(path))
    elif sys.platform == 'darwin': subprocess.Popen(['open', path])
    else: subprocess.Popen(['xdg-open', path])
open_folder(output_path)
