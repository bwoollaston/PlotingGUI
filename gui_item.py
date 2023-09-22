import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd

class UiFrame(tk.Frame):
    def __init__(self, parent, index, file_combo, *args, **kwargs):
        global error_label
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.columnconfigure(0, minsize=120)
        self.columnconfigure(1, minsize=120)
        self.file_combo = file_combo
        self.series_sc = tk.BooleanVar()
        self.sheet_frame = ttk.Frame(self)
        self.sheet_name = ""
        self.x_column_frame = ttk.Frame(self)
        self.y_column_frame = ttk.Frame(self)
        self.sheet_combo = ttk.Combobox(self.sheet_frame, state="readonly")
        self.x_column_combo = ttk.Combobox(self.x_column_frame, state="readonly")
        self.y_column_combo = ttk.Combobox(self.y_column_frame, state="readonly")
        self.initialize(index)
        self.dataFrame = []

    def initialize(self, index):
        # Create widgets and layout for the frame
        self.sheet_frame.columnconfigure(0, minsize=120)
        self.sheet_frame.columnconfigure(1, minsize=120)
        self.x_column_frame.columnconfigure(0, minsize=120)
        self.x_column_frame.columnconfigure(1, minsize=120)
        self.y_column_frame.columnconfigure(0, minsize=120)
        self.y_column_frame.columnconfigure(1, minsize=120)
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

    def set_sheet(self):
        self.sheet_name = self.sheet_combo.get()
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

class Label_entry_item(ttk.Frame):
    def __init__(self,parent,*args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.columnconfigure(0,minsize=120)
        self.columnconfigure(1, minsize=120)
        self.label = ttk.Label(self, text="")
        self.label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry = ttk.Entry(self)
        self.entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
