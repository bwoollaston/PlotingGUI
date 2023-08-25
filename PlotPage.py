import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import configparser

# Global variables
x_axis_column = ""
y_axis_column = ""
sheet_name = ""
plot_canvas = None


# Function to populate sheet names in combo box
def refresh_sheets():
    file_path = file_path_entry.get()

    try:
        excel_sheets = pd.ExcelFile(file_path).sheet_names
        sheet_combo['values'] = excel_sheets
        error_label.config(text="")
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")


# Function to populate column names in combo boxes
def populate_columns(event=None):
    file_path = file_path_entry.get()
    sheet_name = sheet_combo.get()

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        x_column_combo_values = list(df.columns)
        y_column_combo_values = list(df.columns)

        column_combo_x['values'] = x_column_combo_values
        column_combo_y['values'] = y_column_combo_values

        error_label.config(text="")
    except Exception as e:
        error_label.config(text=f"Error: {str(e)}")


# Function to open file explorer and get file path
def browse_file():
    last_used_config = configparser.ConfigParser()
    last_used_config.read("config.ini")
    last_used_file_path = last_used_config.get("LastUsed", "FilePath", fallback="")

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")], initialdir=last_used_file_path)

    if file_path:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_path)
        refresh_sheets()  # Refresh sheet names when a new file is selected
        save_last_used_file(file_path)  # Save the selected file path for next run

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
    x_axis_column = column_combo_x.get()
    y_axis_column = column_combo_y.get()
    sheet_name = sheet_combo.get()

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        plt.figure()
        plt.plot(df[x_axis_column], df[y_axis_column])
        plt.xlabel(x_axis_column)
        plt.ylabel(y_axis_column)
        plt.title(plot_title)

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

sheet_label = ttk.Label(interface_frame, text="Sheet Name:")
sheet_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
sheet_combo = ttk.Combobox(interface_frame, state="readonly")
sheet_combo.grid(row=1, column=1, padx=5, pady=5, columnspan=2)
sheet_combo.bind("<<ComboboxSelected>>", populate_columns)  # Bind event to populate columns

x_column_label = ttk.Label(interface_frame, text="X-Axis Column:")
x_column_label.grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
column_combo_x = ttk.Combobox(interface_frame, state="readonly")
column_combo_x.grid(row=3, column=1, padx=5, pady=5, columnspan=2)

y_column_label = ttk.Label(interface_frame, text="Y-Axis Column:")
y_column_label.grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
column_combo_y = ttk.Combobox(interface_frame, state="readonly")
column_combo_y.grid(row=4, column=1, padx=5, pady=5, columnspan=2)

title_label = ttk.Label(interface_frame, text="Plot Title:")
title_label.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
title_entry = ttk.Entry(interface_frame)
title_entry.grid(row=5, column=1, padx=5, pady=5, columnspan=2)

run_button = ttk.Button(interface_frame, text="Run", command=generate_plot)
run_button.grid(row=6, columnspan=3, padx=5, pady=10)

error_label = ttk.Label(interface_frame, text="", foreground="red")
error_label.grid(row=7, columnspan=3)

# Start the GUI event loop
root.mainloop()