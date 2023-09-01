import os.path
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import configparser
import ClickLabel

# Global variables
x1_axis_column = ""
y1_axis_column = ""
x2_axis_column = ""
y2_axis_column = ""
sheet1_name = ""
sheet2_name = ""
plot_canvas = None



# Function to populate sheet names in combo box
def refresh_sheets():
    file_path = file_path_entry.get()

    try:
        excel_sheets = pd.ExcelFile(file_path).sheet_names
        excel_sheets.insert(0,"")
        sheet1_combo['values'] = excel_sheets
        sheet2_combo['values'] = excel_sheets
        error_label.config(text="")
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")


# Function to populate column names in combo boxes
def populate_columns1(event=None):
    file_path = file_path_entry.get()
    sheet1_name = sheet1_combo.get()
    if sheet1_name == "":
        column_combo_x1['values'] = []
        column_combo_y1['values'] = []
        return
    try:
        df1 = pd.read_excel(file_path, sheet_name=sheet1_name)

        x1_column_combo_values = list(df1.columns)
        y1_column_combo_values = list(df1.columns)

        column_combo_x1['values'] = x1_column_combo_values
        column_combo_y1['values'] = y1_column_combo_values

        error_label.config(text="")
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")
def populate_columns2(event=None):
    file_path = file_path_entry.get()
    sheet2_name = sheet2_combo.get()
    if sheet2_name == "":
        column_combo_x2['values'] = []
        column_combo_y2['values'] = []
        return
    try:
        df2 = pd.read_excel(file_path, sheet_name=sheet2_name)

        x2_column_combo_values = list(df2.columns)
        y2_column_combo_values = list(df2.columns)

        column_combo_x2['values'] = x2_column_combo_values
        column_combo_y2['values'] = y2_column_combo_values

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
    x1_axis_column = column_combo_x1.get()
    y1_axis_column = column_combo_y1.get()
    sheet1_name = sheet1_combo.get()
    sheet2_name = sheet2_combo.get()

    try:
        df1 = pd.read_excel(file_path, sheet_name=sheet1_name)
        plt.figure()
        fig, ax1 = plt.subplots()
        if series1_sc.get():
            ax1.scatter(df1[x1_axis_column], df1[y1_axis_column], label=y1_axis_column)
        else:
            ax1.plot(df1[x1_axis_column], df1[y1_axis_column], label = y1_axis_column)
        ax1.set_xlabel(x1_axis_column)
        ax1.set_ylabel(y1_axis_column)
        if column_combo_x2.get() != "":
            df2 = pd.read_excel(file_path, sheet_name=sheet2_name)
            x2_axis_column = column_combo_x2.get()
            y2_axis_column = column_combo_y2.get()
            if sep_y_axes.get():
                ax2 = ax1.twinx()
                ax2.set_ylabel(y2_axis_column)
                if series2_sc.get():
                    ax2.scatter(df2[x2_axis_column], df2[y2_axis_column], color="red", label=y2_axis_column)
                else:
                    ax2.plot(df2[x2_axis_column], df2[y2_axis_column],color="red",label=y2_axis_column)
                ax2.legend(loc="upper right")
            else:
                if series2_sc.get():
                    ax1.scatter(df2[x2_axis_column], df2[y2_axis_column], color="red", label=y2_axis_column)
                else:
                    ax1.plot(df2[x2_axis_column], df2[y2_axis_column], color="red", label=y2_axis_column)


        ax1.legend(loc="upper left")
        plt.title(plot_title)
        labelHandle = ClickLabel.LabelHandler(ax1)
        plot_canvas = FigureCanvasTkAgg(plt.gcf(), master=plot_frame)
        plot_canvas.draw()
        plot_canvas.get_tk_widget().pack()
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

# Interface elements
file_path_label = ttk.Label(interface_frame, text="File Path:")
file_path_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
file_path_entry = ttk.Entry(interface_frame)
file_path_entry.grid(row=0, column=1, padx=5, pady=5)
browse_button = ttk.Button(interface_frame, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

series1_label = ttk.Label(interface_frame, text="Series 1")
series1_label.grid(row=2,column=0,sticky=tk.W, padx=5, pady=5)

sheet1_label = ttk.Label(interface_frame, text="Sheet Name:")
sheet1_label.grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
sheet1_combo = ttk.Combobox(interface_frame, state="readonly")
sheet1_combo.grid(row=3, column=1, padx=5, pady=5)
sheet1_combo.bind("<<ComboboxSelected>>", populate_columns1)  # Bind event to populate columns

series1_sc = tk.BooleanVar()
series1_scatter_cb = ttk.Checkbutton(interface_frame,text="Scatter",variable=series1_sc)
series1_scatter_cb.grid(row=3, column=2, padx=5, pady=5)

x1_column_label = ttk.Label(interface_frame, text="X-Axis Column:")
x1_column_label.grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
column_combo_x1 = ttk.Combobox(interface_frame, state="readonly")
column_combo_x1.grid(row=4, column=1, padx=5, pady=5)

y1_column_label = ttk.Label(interface_frame, text="Y-Axis Column:")
y1_column_label.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
column_combo_y1 = ttk.Combobox(interface_frame, state="readonly")
column_combo_y1.grid(row=5, column=1, padx=5, pady=5)

series2_label = ttk.Label(interface_frame, text="Series 2")
series2_label.grid(row=6,column=0,sticky=tk.W, padx=5, pady=5)

sheet2_label = ttk.Label(interface_frame, text="Sheet Name:")
sheet2_label.grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
sheet2_combo = ttk.Combobox(interface_frame, state="readonly")
sheet2_combo.grid(row=7, column=1, padx=5, pady=5)
sheet2_combo.bind("<<ComboboxSelected>>", populate_columns2)

series2_sc = tk.BooleanVar()
series2_scatter_cb = ttk.Checkbutton(interface_frame,text="Scatter",variable=series2_sc)
series2_scatter_cb.grid(row=7, column=2, padx=5, pady=5)

x2_column_label = ttk.Label(interface_frame, text="X-Axis Column:")
x2_column_label.grid(row=8, column=0, sticky=tk.W, padx=5, pady=5)
column_combo_x2 = ttk.Combobox(interface_frame, state="readonly")
column_combo_x2.grid(row=8, column=1, padx=5, pady=5)

y2_column_label = ttk.Label(interface_frame, text="Y-Axis Column:")
y2_column_label.grid(row=9, column=0, sticky=tk.W, padx=5, pady=5)
column_combo_y2 = ttk.Combobox(interface_frame, state="readonly")
column_combo_y2.grid(row=9, column=1, padx=5, pady=5)

sep_y_axes = tk.BooleanVar()
seperateAxes_label = ttk.Label(interface_frame, text="Seperate Y-Axes")
seperateAxes_label.grid(row=10, column=0, sticky=tk.W, padx=5, pady=5)
seperateAxes_checkbox = ttk.Checkbutton(interface_frame,variable=sep_y_axes)
seperateAxes_checkbox.grid(row=10, column=1, padx=5, pady=5)

title_label = ttk.Label(interface_frame, text="Plot Title:")
title_label.grid(row=11, column=0, sticky=tk.W, padx=5, pady=5)
title_entry = ttk.Entry(interface_frame)
title_entry.grid(row=11, column=1, padx=5, pady=5, columnspan=2)

run_button = ttk.Button(interface_frame, text="Run", command=generate_plot)
run_button.grid(row=12, columnspan=3, padx=5, pady=10)

error_label = ttk.Label(interface_frame, text="", foreground="red")
error_label.grid(row=13, columnspan=3)

open_file_ini()
# Start the GUI event loop
root.mainloop()