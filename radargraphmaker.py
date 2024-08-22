import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib import rcParams
import mplcursors
import webbrowser
from matplotlib.widgets import Button
from matplotlib.patches import FancyBboxPatch
import matplotlib.colors as mcolors

# Read data from Excel
file_path = 'DataToPlot.xlsx'  # Replace with your Excel file path
df = pd.read_excel(file_path)

# Extract properties and data
properties_df = df.iloc[:12]  # Assuming first 10 rows include new rows for font size, font family, and series names
data_df = df.iloc[12:]

# Extract font size and font family
font_size = int(properties_df[properties_df['Categories'] == 'Font Size'].iloc[0, 3])
font_family = properties_df[properties_df['Categories'] == 'Font Family'].iloc[0, 3]

save_image = properties_df['Save Image'].iloc[0] == 'Yes'
legend_flag = properties_df['Legend'].iloc[0] == 'Yes'


emphasize_series_row = properties_df[properties_df['Categories'] == 'Emphasize Series']
emphasize_series_flags = emphasize_series_row.iloc[0, 3:]

include_series_row = properties_df[properties_df['Categories'] == 'Include']
include_series_flags = include_series_row.iloc[0, 3:]

# Dynamically allocate columns
series_columns = [col for col in properties_df.columns if 'Series' in col]
series_names = [properties_df[properties_df['Categories'] == 'Series Name'][col].values[0] for col in series_columns]

include_mask = data_df['Include'] == True
data_df = data_df[include_mask]

annotate_masks = {
    series: data_df[f'Annotate {i+1}'].values == True
    for i, series in enumerate(series_names)
}

# Set font properties
rcParams['font.family'] = font_family
rcParams['font.size'] = font_size

# Function to invert a value
def _invert(x, limits):
    return limits[1] - (x - limits[0])

# Function to scale data
def _scale_data(data, ranges):
    scaled_data = []
    for d, (y1, y2) in zip(data, ranges):
        if pd.isna(d):
            # Skip N/A values
            scaled_data.append(np.nan)
            continue
        
        if not ((y1 <= d <= y2) or (y2 <= d <= y1)):
            print(f"Warning: Data point {d} is out of the specified range ({y1}, {y2}). Clipping the value.")
            d = max(min(d, max(y1, y2)), min(y1, y2))

        x1, x2 = ranges[0]
        if x1 > x2:
            d = _invert(d, (x1, x2))
            x1, x2 = x2, x1

        sdata = (d - y1) / (y2 - y1) * (x2 - x1) + x1
        scaled_data.append(sdata)

    return scaled_data

# Custom formatter for axis labels
def custom_formatter(x, pos):
    return f'{x:.1f}' if x < 10 else f'{int(x)}'

# Maps for line styles and markers
line_style_map = {
    'solid': '-',
    'dash': '--',
    'dot': ':',
    'dashdot': '-.'
}

marker_map = {
    'circle': 'o',
    'square': 's',
    'triangle': '^',
    'diamond': 'D'
}

# Function to convert rgb(x, y, z) to hex color
def rgb_to_hex(rgb_str):
    rgb = rgb_str.replace('rgb(', '').replace(')', '').split(',')
    return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]), int(rgb[1]), int(rgb[2]))

def lighten_color(color, amount=0.5):
    try:
        c = mcolors.cnames[color]
    except:
        c = color
    c = mcolors.to_rgb(c)
    return mcolors.to_hex([min(1, x + (1 - x) * amount) for x in c])


def adjust_color_for_export(color, emphasize_flag):
    if emphasize_flag:
        return color
    return lighten_color(color, amount=0.8)

class ComplexRadar():
    def __init__(self, fig, variables, ranges, n_ordinate_levels=6, font_size=10):
        angles = np.arange(0, 360, 360./len(variables))
        self.variables = variables
        self.ranges = ranges
        self.series_lines = []

        axes = [fig.add_axes([0.1, 0.1, 0.8, 0.8], polar=True, label=f"axes{i}") for i in range(len(variables))]
        l, text = axes[0].set_thetagrids(angles, labels=variables)
        for txt, angle in zip(text, angles):
            txt.set_rotation(angle-90)
            txt.set_fontsize(font_size)
            txt.set_position((0, -0.1))
        for ax in axes[1:]:
            ax.patch.set_visible(False)
            ax.grid("off")
            ax.xaxis.set_visible(False)
        for i, ax in enumerate(axes):
            grid = np.linspace(*ranges[i], num=n_ordinate_levels)
            gridlabel = [custom_formatter(x, None) for x in grid]
            if ranges[i][0] > ranges[i][1]:
                grid = grid[::-1]
                gridlabel = gridlabel[::-1]
            elif ranges[i][0] < ranges[i][1]:
                gridlabel[0] = ""
            ax.set_rgrids(grid, labels=gridlabel, angle=angles[i])
            ax.set_ylim(*ranges[i])
            for label in ax.get_yticklabels():
                label.set_fontsize(font_size)
                label.set_rotation(45)  # Rotate the label for better readability
                label.set_horizontalalignment('right')  # Align right to avoid overlap
                label.set_verticalalignment('bottom')  # Adjust vertical alignment
        self.angle = np.deg2rad(np.r_[angles, angles[0]])
        self.ax = axes[0]
        self.links = []
        self.series_lines = []  # Store the lines for each series
        self.data_points = []  # Store the data points for hover detection
        self.annotations = []

    def plot(self, data, properties, series_name, annotate_mask, color=None, *args, **kw):
        sdata = _scale_data(data, self.ranges)
        if len(self.angle) != len(sdata) + 1:
            print(f"Error: Angle length ({len(self.angle)}) does not match sdata length ({len(sdata) + 1}) for series {series_name}")
            return
        color = color or rgb_to_hex(properties['Color'])
        linestyle = line_style_map[properties['Line Type']]
        linewidth = properties['Line Thickness']
        marker = marker_map[properties['Marker']]
        markersize = properties['Markersize']
        line, = self.ax.plot(self.angle, np.r_[sdata, sdata[0]], 
                             marker=marker, 
                             markersize=markersize, 
                             linestyle=linestyle, 
                             linewidth=linewidth, 
                             color=color, 
                             label=series_name,
                             *args, **kw)
        if properties['Fill']:
            self.ax.fill(self.angle, np.r_[sdata, sdata[0]], color=color, alpha=0.25)

        self.series_lines.append((line, series_name))

        self.links.append((line, properties['Link']))
        #self.series_lines.append((line, data))  # Store original data points

        # Store the data points with their corresponding angles for hover detection
        for angle, value, orig_value, annotate in zip(self.angle, sdata, data, annotate_mask):
            if not np.isnan(value):
                self.data_points.append((angle, value, line, series_name, data))
                if annotate:
            # add annotations
                    annotation = self.ax.annotate(f'{orig_value:.2f}', xy=(angle, value), xytext=(angle - 0.6, value + 0.5),
                                textcoords='data', ha='center', va='center',
                                bbox=dict(boxstyle='round,pad=0.3', edgecolor='none', facecolor='lightgrey', alpha=0.6),
                                arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0.2', color='grey'))
                    self.annotations.append(annotation)
    
    def clear_annotations(self):
        for annotation in self.annotations:
            annotation.remove()
        self.annotations.clear()
    
    def get_original_value(self, angle, data):
        index = np.argmin(np.abs(self.angle - angle))
        return self.variables[index], data[index]

# Convert properties to a dictionary for each series
properties = {}
for col in properties_df.columns[3:]:
    series_name = properties_df[properties_df['Categories'] == 'Series Name'][col].values[0]
    if pd.notna(series_name):
        properties[series_name] = {
            'Color': properties_df[properties_df['Categories'] == 'Color'][col].values[0],
            'Line Thickness': int(properties_df[properties_df['Categories'] == 'Line Thickness'][col].values[0]) if isinstance(properties_df[properties_df['Categories'] == 'Line Thickness'][col].values[0], (int, str)) and str(properties_df[properties_df['Categories'] == 'Line Thickness'][col].values[0]).isdigit() else 1,
            'Line Type': properties_df[properties_df['Categories'] == 'Line Type'][col].values[0],
            'Marker': properties_df[properties_df['Categories'] == 'Marker'][col].values[0],
            'Markersize': int(properties_df[properties_df['Categories'] == 'Markersize'][col].values[0]) if isinstance(properties_df[properties_df['Categories'] == 'Markersize'][col].values[0], (int, str)) and str(properties_df[properties_df['Categories'] == 'Markersize'][col].values[0]).isdigit() else 1,
            'Fill': properties_df[properties_df['Categories'] == 'Fill'][col].values[0] == 'True',
            'Link': properties_df[properties_df['Categories'] == 'Link'][col].values[0]
        }

# Extract categories and ranges
categories = list(data_df['Categories'])
ranges = [(data_df['Range_min'].iloc[i], data_df['Range_max'].iloc[i]) for i in range(len(data_df))]

# Transpose data for easier plotting
data = data_df.drop(['Categories', 'Range_min', 'Range_max', 'Include', 'Annotate 1', 'Annotate 2', 'Annotate 3'], axis=1).T.values

if len(data) > len(series_names):
    data = data[:len(series_names)]

print(f"Data shape: {data.shape}")
print(f"Number of series: {len(series_names)}")

# Plotting
fig1 = plt.figure(figsize=(16, 12))  # Set figure size to 1600 x 1200 pixels
radar = ComplexRadar(fig1, categories, ranges, font_size=font_size)

radar.clear_annotations()

# Plot each series
for i, series in enumerate(data):
    series_name = series_names[i]
    include_flag = include_series_flags.iloc[i]
    if not include_flag:
        continue
    print(f"Plotting series: {series_name}")
    radar.plot(series, properties[series_name], series_name, annotate_masks[series_name])

# Add legend
#plt.legend(loc='upper right', bbox_to_anchor=(1.1, 1.1))

# Interactive cursor to open links on click
cursor = mplcursors.cursor(radar.ax, hover=True)

# Customize the annotation box to show the axis value
@cursor.connect("add")
def on_add(sel):
    angle = sel.target[0]
    r_value = sel.target[1]
    closest_point = None
    min_distance = float('inf')

    # Find the closest data point
    for point_angle, point_r, line, series_name, data in radar.data_points:
        distance = np.sqrt((angle - point_angle) ** 2 + (r_value - point_r) ** 2)
        if distance < min_distance:
            min_distance = distance
            closest_point = (point_angle, point_r, line, series_name, data)

    if closest_point and min_distance < 0.2:  # Increase the threshold for sensitivity
        variable, value = radar.get_original_value(closest_point[0], closest_point[4])
        sel.annotation.set_text(f'{closest_point[3]}\n{variable}: {value:.2f}')  # Display series name and axis value
        sel.annotation.set_visible(True)
        for line, _ in radar.series_lines:
            if sel.artist == line:
                line.set_alpha(1.0)  # Emphasize the hovered series
                line.set_linewidth(3)  # Make the line thicker
            else:
                line.set_alpha(0.2)  # De-emphasize all other series
                line.set_linewidth(1)  # Reset line width
    else:
        sel.annotation.set_visible(False)  # Hide the annotation if not near a datapoint
    fig1.canvas.draw_idle()  # Update the plot

def on_leave(event):
    for line, _ in radar.series_lines:
        line.set_alpha(1.0)  # Reset emphasis
        line.set_linewidth(1.5)  # Reset line width
    fig1.canvas.draw_idle()  # Update the plot

def on_click(event):
    for line, link in radar.links:
        if line.contains(event)[0]:
            webbrowser.open(link)
            break

fig1.canvas.mpl_connect("figure_leave_event", on_leave)
fig1.canvas.mpl_connect("button_press_event", on_click)

if legend_flag:
    legend_handles, legend_labels = zip(*[(line, series_name) for line, series_name in radar.series_lines])
    plt.legend(handles=legend_handles, labels=legend_labels, loc='upper right', bbox_to_anchor=(1.1, 1.1))
#plt.legend(loc='upper right', bbox_to_anchor=(1.1, 1.1))  # Adjust the legend position
plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)  # Adjust subplot parameters for more white space
plt.show()


# save image
if save_image:
    fig2 = plt.figure(figsize=(16, 12))
    newradar = ComplexRadar(fig2, categories, ranges, font_size=font_size)
    for i, series in enumerate(data):
        series_name = series_names[i]
        include_flag = include_series_flags.iloc[i]

        if not include_flag:
            continue

        emphasize_flag = emphasize_series_flags.iloc[i]
        emphasize_flag = emphasize_flag == True
        print(f"Series: {series_name}, Emphasize: {emphasize_flag}")

        color = adjust_color_for_export(rgb_to_hex(properties[series_name]['Color']), emphasize_flag)
        newradar.plot(series, properties[series_name], series_name, annotate_masks[series_name], color=color)
    
    output_path = './radar_chart.jpg'
    plt.savefig(output_path, format='jpg', bbox_inches='tight')
    plt.close(fig2)
#plt.show()




















