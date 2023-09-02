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

# Global variables
Debug = True
series_count = 1
series_list = []
x_new = []
y_new = []
plot_canvas = None
fig, ax1 = plt.subplots()
ax2 = ax1

# Function to populate sheet names in combo box
def refresh_sheets():
    file_path = file_path_entry.get()

    try:
        excel_sheets = pd.ExcelFile(file_path).sheet_names
        excel_sheets.insert(0,"")
        series1.sheet_combo['values'] = excel_sheets
        series2.sheet_combo['values'] = excel_sheets
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
    global plot_canvas

    # Clear previous plot
    if plot_canvas:
        plot_canvas.get_tk_widget().pack_forget()

    file_path = file_path_entry.get()
    plot_title = title_entry.get()
    x_axis_column = series1.x_column_combo.get()
    y_axis_column = series1.y_column_combo.get()
    sheet1_name = series1.sheet_combo.get()
    sheet2_name = series2.sheet_combo.get()

    try:
        series1.dataFrame = pd.read_excel(file_path, sheet_name=sheet1_name)
        plt.figure()
#        fig, ax1 = plt.subplots()
        if series1.series_sc.get():
            ax1.scatter(series1.dataFrame[x_axis_column], series1.dataFrame[y_axis_column], label=y_axis_column)
        else:
            ax1.plot(series1.dataFrame[x_axis_column], series1.dataFrame[y_axis_column], label = y_axis_column)
        ax1.set_xlabel(x_axis_column)
        ax1.set_ylabel(y_axis_column)
        if series2.x_column_combo.get() != "":
            series2.dataFrame = pd.read_excel(file_path, sheet_name=sheet2_name)
            x_axis_column = series2.x_column_combo.get()
            y_axis_column = series2.y_column_combo.get()
            if sep_y_axes.get():
                ax2 = ax1.twinx()
                ax2.set_ylabel(y_axis_column)
                if series2.series_sc.get():
                    ax2.scatter(series2.dataFrame[x_axis_column], series2.dataFrame[y_axis_column], color="red", label=y_axis_column)
                else:
                    ax2.plot(series2.dataFrame[x_axis_column], series2.dataFrame[y_axis_column],color="red",label=y_axis_column)
                ax2.legend(loc="upper right")
            else:
                if series2.series_sc.get():
                    ax1.scatter(series2.dataFrame[x_axis_column], series2.dataFrame[y_axis_column], color="red", label=y_axis_column)
                else:
                    ax1.plot(series2.dataFrame[x_axis_column], series2.dataFrame[y_axis_column], color="red", label=y_axis_column)
        if len(x_new) == 0:
            ax1.plot(x_new, y_new, color="green", label="Fit")

        ax1.legend(loc="upper left")
        plt.title(plot_title)
        #labelHandle = ClickLabel.LabelHandler(ax1)
        plot_canvas.draw()
        plot_canvas.get_tk_widget().pack()
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")


def fitSeriesData():
    x_pts = []
    y_pts = []
    file_path = file_path_entry.get()
    sheet1_name = series1.sheet_combo.get()
    series1.dataFrame = pd.read_excel(file_path, sheet_name=sheet1_name)
    x_column = series1.x_column_combo.get()
    y_column = series1.y_column_combo.get()
    x = series1.dataFrame[x_column]
    y = series1.dataFrame[y_column]
    for item in x:
        x_pts.append(item)
    for item in y:
        y_pts.append(item)
    pyFit = np.polyfit(x_pts, y_pts, int(degree_input.get()))
    pyfit = np.poly1d(pyFit)
    x_new = np.linspace(x_pts[0], x_pts[-1], 50)
    y_new = pyfit(x_new)
    ax1.set_xlabel("X Axis")
    ax1.set_ylabel("Y Axis")
    ax1.set_title("Polynomial Fit to Data")
    ax1.plot(x_new, y_new, color="green", label="Fit")
    remove_plt()
    ax1.legend()
    plot_canvas.draw()

def remove_plt():
    i = 0
    for line in ax1.get_lines():
        ax1.get_lines()[i].remove()
        i += 1
    i = 0
    for line in ax2.get_lines():
        ax2.get_lines()[i].remove()
        i += 1

class UiFrame(tk.Frame):
    def __init__(self, parent, index, file_combo, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.file_combo = file_combo
        self.series_sc = tk.BooleanVar()
        self.sheet_frame = ttk.Frame(self)
        self.x_column_frame = ttk.Frame(self)
        self.y_column_frame = ttk.Frame(self)
        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.x_column_combo = ttk.Combobox(self.x_column_frame, state="readonly")
        self.y_column_combo = ttk.Combobox(self.y_column_frame, state="readonly")
        self.initialize(index)
        self.dataFrame = []

    def initialize(self, index):
        # Create widgets and layout for the frame
        series_label = ttk.Label(self, text="Series " + str(index))
        series_label.pack(padx=5, pady=5)

        self.sheet_frame.pack()
        sheet_label = ttk.Label(self.sheet_frame, text="Sheet Name:")
        sheet_label.grid(row=0, column=0, padx=5, pady=5)

        self.sheet_combo.grid(row=0, column=1, padx=5, pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.populate_columns)  # Bind event to populate columns
        scatter_cb = ttk.Checkbutton(self.sheet_frame, text="Scatter", variable=self.series_sc)
        scatter_cb.grid(row=0, column=2, padx=5, pady=5)

        self.x_column_frame.pack()
        x_column_label = ttk.Label(self.x_column_frame,text="X-Axis Column:")
        x_column_label.grid(row=0, column=0, padx=5, pady=5)

        self.x_column_combo.grid(row=0, column=1, columnspan=2, padx=5, pady=5)

        self.y_column_frame.pack()
        y_column_label = ttk.Label(self.y_column_frame,text="Y-Axis Column:")
        y_column_label.grid(row=0, column=0, padx=5, pady=5)
        self.y_column_combo.grid(row=0, column=1, padx=5, pady=5)

    def populate_columns(self, event=None):
        file_path = self.file_combo.get()
        sheet_name = self.sheet_combo.get()
        if sheet_name == "":
            self.x_column_combo['values'] = []
            self.y_column_combo['values'] = []
            return
        try:
            self.dataFrame = pd.read_excel(file_path, sheet_name=sheet_name)
            x_column_combo_values = list(self.dataFrame.columns)
            y_column_combo_values = list(self.dataFrame.columns)
            self.x_column_combo['values'] = x_column_combo_values
            self.y_column_combo['values'] = y_column_combo_values
            error_label.config(text="")
        except Exception as e:
            error_label.config(text=f"Error: {str(e)}")

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

series1 = UiFrame(interface_frame, index="1", file_combo=file_path_entry)
series1.pack()
series2 = UiFrame(interface_frame, index="2", file_combo=file_path_entry)
series2.pack()
series_list.append(series1)
series_list.append(series2)

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
fit_button = ttk.Button(button_frame, text="Fit", command=fitSeriesData)
fit_button.grid(row=0, column = 0, padx=5, pady=10)
run_button = ttk.Button(button_frame, text="Run", command=generate_plot)
run_button.grid(row=0, column=1, padx=5, pady=10)

error_label = ttk.Label(interface_frame, text="", foreground="red")
error_label.pack()

# Fit elements
degree_label = tk.Label(fit_frame,text="Polynomial Degree")
degree_label.grid(column=0,row=1,sticky=tk.W,padx=5,pady=5)
degree_input = ttk.Entry(fit_frame)
degree_input.grid(column=1,row=1,sticky=tk.W,padx=5,pady=5)
plot_canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)

sv_ttk.set_theme("dark")

open_file_ini()
if Debug:
    series1.sheet_combo.set('Known_Light')
    series1.x_column_combo.set("Index of Refraction")
    series1.y_column_combo.set("Wavelength [nm]")
# Start the GUI event loop
root.mainloop()