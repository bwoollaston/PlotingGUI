import os.path
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import configparser
import ClickLabel
import numpy as np
import sv_ttk
import textwrap
import gui_item as ui

# Global variables
Debug = False
x_axis_column = ""
y_axis_column = ""
series_count = 1
series_list = []
line_list = []
x_new = []
y_new = []
plot_canvas = None
fig, ax1 = plt.subplots()
ax2 = ax1

line_fit = None
seriesFit = np.poly1d

# Function to populate sheet names in combo box
def refresh_sheets():
    file_path = file_path_entry.get()

    try:
        excel_sheets = pd.ExcelFile(file_path).sheet_names
        excel_sheets.insert(0,"")
        for item in series_list:
            item.sheet_combo['values'] = excel_sheets
        error_label.config(text="")
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")

# Function to open file explorer and get file path
def browse_file():
    last_used_config = configparser.ConfigParser()
    last_used_config.read("config.ini")
    last_used_file_path = last_used_config.get("LastUsed", "FilePath", fallback="")
    last_used_dir = os.path.dirname(last_used_file_path)
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")], initialdir=last_used_dir)

    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)
        refresh_sheets()  # Refresh sheet names when a new file is selected
        save_last_used_file(file_path)  # Save the selected file path for next run

def open_file_ini():
    last_used_config = configparser.ConfigParser()
    last_used_config.read("config.ini")
    last_used_file_path = last_used_config.get("LastUsed", "FilePath", fallback="")
    if last_used_file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, last_used_file_path)
        refresh_sheets()  # Refresh sheet names when a new file is selected

def save_last_used_file(file_path):
    config = configparser.ConfigParser()
    config.read("config.ini")
    config.set("LastUsed", "FilePath", file_path)

    with open("config.ini", "w") as config_file:
        config.write(config_file)

# Function to generate the plot
def generate_plot():
    global plot_canvas, x_axis_column, y_axis_column, line_list, line_fit
    # Clear previous plot
    if plot_canvas:
        plot_canvas.get_tk_widget().pack_forget()

    file_path = file_path_entry.get()
    plot_title = title_entry.get()
    plot_title = textwrap.fill(plot_title, width=40)

    for item in series_list:
        item.set_sheet()
    if len(line_list) != 0:
        for item in line_list:
            if item != None:
                item.remove()
    if line_fit != None:
        line_fit.remove()

    try:
        for item in series_list:
            if item.sheet_name != "":
                item.dataFrame = pd.read_excel(file_path, sheet_name=item.sheet_name)
        plt.figure()
        i=0
        for item in series_list:
            if item.x_column_combo.get() != "":
                x_axis_column = series_list[i].x_column_combo.get()
                y_axis_column = series_list[i].y_column_combo.get()
                if item.series_sc.get():
                    line, = ax1.scatter(item.dataFrame[x_axis_column], item.dataFrame[y_axis_column], label=f"{y_axis_column} vs.\n{x_axis_column}")
                    line_list.append(line)
                else:
                    line, = ax1.plot(item.dataFrame[x_axis_column], item.dataFrame[y_axis_column], label=f"{y_axis_column} vs.\n{x_axis_column}")
                    line_list.append(line)
            i += 1
        ax1.set_xlabel(x_axis_column)
        ax1.set_ylabel(y_axis_column)
        # if len(x_new) == 0:
        #     line_fit, = ax1.plot(x_new, y_new, color="green", label="Fit")

        legend1 = ax1.legend(loc="upper right")
        fig.suptitle(plot_title)
        plot_canvas.draw()
        plot_canvas.get_tk_widget().pack()
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")


def fitSeriesData():
    global seriesFit, line_fit
    current_series = int(series_fit_input.entry.get()) - 1
    if line_fit != None:
        line_fit.remove()
    x_pts = []
    y_pts = []
    file_path = file_path_entry.get()
    sheet1_name = series_list[current_series].sheet_combo.get()
    series_list[current_series].dataFrame = pd.read_excel(file_path, sheet_name=sheet1_name)
    x_column = series_list[current_series].x_column_combo.get()
    y_column = series_list[current_series].y_column_combo.get()
    x = series_list[current_series].dataFrame[x_column]
    y = series_list[current_series].dataFrame[y_column]
    for item in x:
        x_pts.append(item)
    for item in y:
        y_pts.append(item)
    pyFit = np.polyfit(x_pts, y_pts, int(degree_input.entry.get()))
    seriesFit = np.poly1d(pyFit)
    x_new = np.linspace(x_pts[0], x_pts[-1], 50)
    y_new = seriesFit(x_new)
    display_polynomial()
    line_fit, = ax1.plot(x_new, y_new, label=f"Fit {y_axis_column} vs.\n{x_axis_column}")
    ax1.legend()
    plot_canvas.draw()

def evaluate_y():
    global seriesFit
    x = float(x_input.entry.get())
    y = seriesFit(x)
    y_input.entry.delete(0, "end")
    y_input.entry.insert(0, str(y))
def evaluate_x():
    y = y_input.entry.get()
    x = y
    x_input.entry.delete(0, "end")
    x_input.entry.insert(0, str(x))

def display_polynomial():
    global seriesFit
    s = ""
    i = 0
    j = len(seriesFit.c) - 1
    for coef in seriesFit.c:
        if j>0:
            s += f"{coef}x^{j} + "
        else:
            s += f"{coef}"
        j -= 1
    polynomial_label.config(text=s)

def add_series():
    for item in series_list:
        item.pack_forget()
    series = ui.UiFrame(series_frame, index=str(len(series_list)+1), file_combo=file_path_entry)
    series_list.append(series)
    refresh_sheets()
    for item in series_list:
        item.pack()
def delete_series():
    for item in series_list:
        item.pack_forget()
    series_list.remove(series_list[len(series_list)-1])
    for item in series_list:
        item.pack()


# Create the main window
root = tk.Tk()
root.title("Matplotlib GUI")

# Create plot frame
plot_frame = ttk.Frame(root)
plot_frame.pack(side=tk.LEFT, padx=10, pady=10)

# Create interface frame
interface_frame = ttk.Frame(root)
interface_frame.pack(side=tk.LEFT, padx=10, pady=10)
fp_frame = ttk.Frame(interface_frame)
fp_frame.pack()

# Create Fit Output Frame
fit_frame = ttk.Frame(root)
fit_frame.pack(side=tk.LEFT)

# Interface elements
file_path_label = ttk.Label(fp_frame, text="File Path:")
file_path_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
file_path_entry = ttk.Entry(fp_frame)
file_path_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
browse_button = ttk.Button(fp_frame, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

series_frame = ttk.Frame(interface_frame)
series_frame.pack()

error_label = ttk.Label(interface_frame, text="", foreground="red")
add_series()
add_series()

sep_y_axes = tk.BooleanVar()
separateAxes_checkbox = ttk.Checkbutton(interface_frame,variable=sep_y_axes, text="Separate Y-Axes")
separateAxes_checkbox.pack(padx=5, pady=5)

title_frame = ttk.Frame(interface_frame)
title_frame.pack()
title_label = ttk.Label(title_frame, text="Plot Title:")
title_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
title_entry = ttk.Entry(title_frame)
title_entry.grid(row=0, column=1, padx=5, pady=5, columnspan=2)

button_frame = ttk.Frame(interface_frame)
button_frame.pack()

add_series_button = ttk.Button(button_frame, text="Add Series", command=add_series)
add_series_button.grid(row=0, column=0, padx=5, pady=5)
delete_series_button = ttk.Button(button_frame, text="Delete Series", command=delete_series)
delete_series_button.grid(row=0, column=1, padx=5, pady=5)

run_button = ttk.Button(interface_frame, text="Run", command=generate_plot)
run_button.pack()


error_label.pack()

# Fit elements
polynomial_label = ttk.Label(fit_frame)
polynomial_label.pack()

series_fit_input = ui.Label_entry_item(fit_frame)
series_fit_input.label.config(text="Series to Fit")
series_fit_input.pack()

degree_input = ui.Label_entry_item(fit_frame)
degree_input.label.config(text="Polynomial Degree")
degree_input.pack()
fit_button = ttk.Button(fit_frame, text="Fit", command=fitSeriesData)
fit_button.pack()

plot_canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)

x_input = ui.Label_entry_item(fit_frame)
x_input.label.config(text="X Value:")
x_input.pack()
evalx_button = ttk.Button(fit_frame, command=evaluate_y, text="Solve f(x)")
evalx_button.pack()

y_input = ui.Label_entry_item(fit_frame)
y_input.label.config(text="Y Value:")
y_input.pack()
evaly_button = ttk.Button(fit_frame, command=evaluate_x, text="Solve x")
evaly_button.pack()

sv_ttk.set_theme("dark")

open_file_ini()
if Debug:
    series_list[0].sheet_combo.set('Known_Light')
    series_list[0].x_column_combo.set("Index of Refraction")
    series_list[0].y_column_combo.set("Wavelength [nm]")
# Start the GUI event loop
root.mainloop()