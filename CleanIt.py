import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import sys
import json
from datetime import datetime

class CleanItApp:
    CONFIG_FILE = os.path.join(os.path.expanduser('~'), '.cleanit_config.json')

    def __init__(self, root):
        self.root = root
        self.root.title("CleanIt")

        # --- Fullscreen/Maximized Window ---
        try:
            self.root.state('zoomed')
        except tk.TclError:
            pass

        self.root.resizable(True, True)

        # --- Colors ---
        self.bg_color = "#E6F7FF"
        self.fg_color = "#000080"

        self.root.configure(bg=self.bg_color)

        # --- Styles for ttk widgets ---
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, font=('Helvetica', 10))
        self.style.configure('TButton', background='#A3D9F7', foreground=self.fg_color, font=('Helvetica', 10, 'bold'), padding=5)
        self.style.map('TButton',
                       background=[('active', '#8BCDF2')],
                       foreground=[('active', self.fg_color)])
        self.style.configure('TEntry', fieldbackground='white', foreground='black', font=('Helvetica', 10))
        self.style.configure('TCombobox', fieldbackground='white', foreground='black', font=('Helvetica', 10))
        self.style.map('TCombobox',
                       fieldbackground=[('readonly', 'white')],
                       selectbackground=[('readonly', 'white')],
                       selectforeground=[('readonly', 'black')])
        self.style.configure('TRadiobutton', background=self.bg_color, foreground=self.fg_color)
        self.style.configure('TCheckbutton', background=self.bg_color, foreground=self.fg_color)
        self.style.configure('TNotebook', background=self.bg_color, borderwidth=0)
        self.style.configure('TNotebook.Tab', background='#C2E5FF', foreground=self.fg_color, font=('Helvetica', 10, 'bold'))
        self.style.map('TNotebook.Tab',
                       background=[('selected', self.bg_color), ('active', '#A3D9F7')],
                       foreground=[('selected', self.fg_color)])

        # Treeview style for missing values
        self.style.configure("Treeview",
                             background="white",
                             foreground="black",
                             rowheight=25,
                             fieldbackground="white")
        self.style.map('Treeview',
                       background=[('selected', self.fg_color)],
                       foreground=[('selected', 'white')])
        self.style.configure("Treeview.Heading",
                             font=('Helvetica', 10, 'bold'),
                             background='#A3D9F7',
                             foreground=self.fg_color)


        # --- Variables ---
        self.input_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()

        # The DataFrame itself, now a class attribute
        self.df = None

        # Date format options
        self._date_format_options = [
            ("MM/DD/YYYY (e.g., 01/15/2023)", "%m/%d/%Y"),
            ("Weekday, Month DD, YYYY (e.g., Sunday, January 15, 2023)", "%A, %B %d, %Y"),
            ("YYYY-MM-DD (e.g., 2023-01-15)", "%Y-%m-%d"),
            ("MM/DD (e.g., 01/15)", "%m/%d"),
            ("MM/DD/YY (e.g., 01/15/23)", "%m/%d/%y"),
            ("DD-Mon (e.g., 15-Jan)", "%d-%b"),
            ("DD-Mon-YY (e.g., 15-Jan-23)", "%d-%b-%y"),
            ("YY-Mon-DD (e.g., 23-Jan-15)", "%y-%b-%d"),
            ("Mon-DD (e.g., Jan-15)", "%b-%d"),
            ("Month-DD (e.g., January-15)", "%B-%d"),
            ("Month DD, YYYY (e.g., January 15, 2023)", "%B %d, %Y"),
            ("YYYY-MM-DD HH:MM:SS (24-hour, e.g., 2023-01-15 13:30:45)", "%Y-%m-%d %H:%M:%S"),
            ("MM/DD/YYYY HH:MM AM/PM (e.g., 01/15/2023 01:30 PM)", "%m/%d/%Y %I:%M %p")
        ]
        self.date_format_display_strings = [item[0] for item in self._date_format_options]
        self._date_format_map = {item[0]: item[1] for item in self._date_format_options}
        self.selected_date_format_display = tk.StringVar(value=self.date_format_display_strings[0])

        # Sorting variables
        self.sort_column = tk.StringVar()
        self.sort_order = tk.BooleanVar(value=True) # True for Ascending, False for Descending

        # Cleaning Options Checkbox Variables
        self.do_trim_whitespace = tk.BooleanVar(value=True)
        self.do_capitalize_strings = tk.BooleanVar(value=True)
        self.do_remove_duplicates = tk.BooleanVar(value=True)


        # Dynamic column checkboxes
        self.date_column_vars = {} # Stores {column_name: tk.BooleanVar} for date formatting
        self.duplicate_check_column_vars = {} # Stores {column_name: tk.BooleanVar} for duplicate detection

        # Missing values Treeview related
        self.missing_values_tree = None # Will be initialized in create_widgets
        self.missing_values_data = {} # Stores original DataFrame index to Treeview iid mapping for editing

        self._load_config() # Load paths on startup
        self.create_widgets()

    def _load_config(self):
        """Loads last used paths from a config file."""
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.input_file_path.set(config.get('last_input_file_path', ''))
                    self.output_folder_path.set(config.get('last_output_folder', ''))
            except Exception as e:
                messagebox.showwarning("Config Load Error", f"Could not load configuration: {e}")

    def _save_config(self):
        """Saves current paths to a config file."""
        config = {
            'last_input_file_path': self.input_file_path.get(),
            'last_output_folder': self.output_folder_path.get()
        }
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            messagebox.showwarning("Config Save Error", f"Could not save configuration: {e}")

    def create_widgets(self):
        # --- Main Frame ---
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(expand=True, fill='both')

        # --- Title ---
        title_label = ttk.Label(main_frame, text="CleanIt: Data Cleaning & Formatting",
                                font=('Helvetica', 16, 'bold'), foreground=self.fg_color)
        title_label.pack(pady=(0, 20), fill='x')

        # --- Notebook (Tabs) ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(expand=True, fill='both')

        # --- Tab 1: File & Output ---
        file_output_tab = ttk.Frame(self.notebook, padding="15 15 15 15")
        self.notebook.add(file_output_tab, text="1. File & Output")

        file_output_tab.grid_columnconfigure(1, weight=1)
        file_output_tab.grid_rowconfigure(0, weight=1)
        file_output_tab.grid_rowconfigure(1, weight=1)

        ttk.Label(file_output_tab, text="Input File:").grid(row=0, column=0, sticky='w', pady=5)
        self.input_entry = ttk.Entry(file_output_tab, textvariable=self.input_file_path, state='readonly')
        self.input_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(file_output_tab, text="Browse...", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(file_output_tab, text="Output Folder:").grid(row=1, column=0, sticky='w', pady=5)
        self.output_entry = ttk.Entry(file_output_tab, textvariable=self.output_folder_path, state='readonly')
        self.output_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(file_output_tab, text="Browse...", command=self.browse_output_folder).grid(row=1, column=2, padx=5, pady=5)


        # --- Tab 2: Cleaning Options ---
        cleaning_options_tab = ttk.Frame(self.notebook, padding="15 15 15 15")
        self.notebook.add(cleaning_options_tab, text="2. Cleaning Options")
        cleaning_options_tab.grid_columnconfigure(0, weight=1)

        row_counter = 0

        # Whitespace Trimming
        ttk.Checkbutton(cleaning_options_tab, text="Trim Leading/Trailing Whitespace", variable=self.do_trim_whitespace).grid(row=row_counter, column=0, sticky='w', pady=(0,2))
        row_counter += 1
        ttk.Label(cleaning_options_tab, text="  - Removes any spaces at the beginning or end of text values in all columns.",
                  font=('Helvetica', 9, 'italic'), foreground='gray', wraplength=500, justify='left').grid(row=row_counter, column=0, sticky='w', padx=10, pady=(0,10))
        row_counter += 1

        # Capitalization
        ttk.Checkbutton(cleaning_options_tab, text="Capitalize Text Fields (Title Case)", variable=self.do_capitalize_strings).grid(row=row_counter, column=0, sticky='w', pady=(10,2))
        row_counter += 1
        ttk.Label(cleaning_options_tab, text="  - Converts the first letter of each word in text fields to uppercase (e.g., 'john doe' -> 'John Doe').",
                  font=('Helvetica', 9, 'italic'), foreground='gray', wraplength=500, justify='left').grid(row=row_counter, column=0, sticky='w', padx=10, pady=(0,10))
        row_counter += 1
        
        # Duplicate Handling
        ttk.Checkbutton(cleaning_options_tab, text="Remove Duplicate Rows", variable=self.do_remove_duplicates).grid(row=row_counter, column=0, sticky='w', pady=(10,2))
        row_counter += 1
        ttk.Label(cleaning_options_tab, text="  - Identifies and removes exact duplicate rows, keeping the first occurrence.",
                  font=('Helvetica', 9, 'italic'), foreground='gray', wraplength=500, justify='left').grid(row=row_counter, column=0, sticky='w', padx=10, pady=(0,5))
        row_counter += 1
        ttk.Label(cleaning_options_tab, text="Select columns to check for duplicates (all must match):",
                  font=('Helvetica', 10, 'bold'), foreground=self.fg_color).grid(row=row_counter, column=0, sticky='w', pady=(5,5))
        row_counter += 1

        # Frame for dynamic duplicate check column checkboxes with scrollbar
        self.duplicate_check_columns_outer_frame = ttk.Frame(cleaning_options_tab)
        self.duplicate_check_columns_outer_frame.grid(row=row_counter, column=0, sticky='nsew', padx=5, pady=5)
        self.duplicate_check_columns_outer_frame.grid_columnconfigure(0, weight=1)
        self.duplicate_check_columns_outer_frame.grid_rowconfigure(0, weight=1)

        self.duplicate_check_columns_canvas = tk.Canvas(self.duplicate_check_columns_outer_frame, background=self.bg_color, highlightthickness=0, height=100) # Fixed height
        self.duplicate_check_columns_canvas.grid(row=0, column=0, sticky='nsew')

        self.duplicate_check_columns_scrollbar = ttk.Scrollbar(self.duplicate_check_columns_outer_frame, orient="vertical", command=self.duplicate_check_columns_canvas.yview)
        self.duplicate_check_columns_scrollbar.grid(row=0, column=1, sticky='ns')

        self.duplicate_check_columns_canvas.configure(yscrollcommand=self.duplicate_check_columns_scrollbar.set)
        self.duplicate_check_columns_canvas.bind('<Configure>', lambda e: self.duplicate_check_columns_canvas.configure(scrollregion = self.duplicate_check_columns_canvas.bbox("all")))
        self.duplicate_check_columns_canvas.bind_all("<MouseWheel>", self._on_mousewheel_duplicate_check)

        self.duplicate_check_columns_frame = ttk.Frame(self.duplicate_check_columns_canvas, padding="5 5 5 5")
        self.duplicate_check_columns_canvas.create_window((0, 0), window=self.duplicate_check_columns_frame, anchor="nw")
        self.duplicate_check_columns_frame.bind("<Configure>", lambda e: self.duplicate_check_columns_canvas.configure(scrollregion=self.duplicate_check_columns_canvas.bbox("all")))

        ttk.Label(self.duplicate_check_columns_frame, text="Load a file to see columns...", background=self.bg_color).pack(pady=10)


        # --- Tab 3: Date Formatting --- (MOVED FROM 4)
        date_formatting_tab = ttk.Frame(self.notebook, padding="15 15 15 15")
        self.notebook.add(date_formatting_tab, text="3. Date Formatting") # Updated tab number

        date_formatting_tab.grid_columnconfigure(1, weight=1)
        date_formatting_tab.grid_rowconfigure(2, weight=1)

        ttk.Label(date_formatting_tab, text="Choose Output Date Format:").grid(row=0, column=0, sticky='w', pady=5)
        self.date_format_combobox = ttk.Combobox(date_formatting_tab, textvariable=self.selected_date_format_display,
                                                 values=self.date_format_display_strings, state='readonly')
        self.date_format_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        if self.date_format_display_strings:
            self.selected_date_format_display.set(self.date_format_display_strings[0])

        ttk.Label(date_formatting_tab, text="Select Columns to Format as Dates:").grid(row=1, column=0, sticky='nw', pady=5)
        
        self.date_columns_outer_frame = ttk.Frame(date_formatting_tab)
        self.date_columns_outer_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='nsew')
        self.date_columns_outer_frame.grid_columnconfigure(0, weight=1)
        self.date_columns_outer_frame.grid_rowconfigure(0, weight=1)

        self.date_columns_canvas = tk.Canvas(self.date_columns_outer_frame, background=self.bg_color, highlightthickness=0)
        self.date_columns_canvas.grid(row=0, column=0, sticky='nsew')

        self.date_columns_scrollbar = ttk.Scrollbar(self.date_columns_outer_frame, orient="vertical", command=self.date_columns_canvas.yview)
        self.date_columns_scrollbar.grid(row=0, column=1, sticky='ns')

        self.date_columns_canvas.configure(yscrollcommand=self.date_columns_scrollbar.set)
        self.date_columns_canvas.bind('<Configure>', lambda e: self.date_columns_canvas.configure(scrollregion = self.date_columns_canvas.bbox("all")))
        self.date_columns_canvas.bind_all("<MouseWheel>", self._on_mousewheel_date_format)

        self.date_columns_frame = ttk.Frame(self.date_columns_canvas, padding="5 5 5 5")
        self.date_columns_canvas.create_window((0, 0), window=self.date_columns_frame, anchor="nw")
        self.date_columns_frame.bind("<Configure>", lambda e: self.date_columns_canvas.configure(scrollregion=self.date_columns_canvas.bbox("all")))

        ttk.Label(self.date_columns_frame, text="Load a file to see columns...", background=self.bg_color).pack(pady=10)


        # --- Tab 4: Sorting --- (MOVED FROM 5)
        sorting_tab = ttk.Frame(self.notebook, padding="15 15 15 15")
        self.notebook.add(sorting_tab, text="4. Sorting") # Updated tab number

        sorting_tab.grid_columnconfigure(1, weight=1)
        sorting_tab.grid_rowconfigure(0, weight=1)

        ttk.Label(sorting_tab, text="Sort Data By Column:").grid(row=0, column=0, sticky='w', pady=5)
        self.sort_column_combobox = ttk.Combobox(sorting_tab, textvariable=self.sort_column,
                                                 values=[], state='readonly')
        self.sort_column_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        ttk.Label(sorting_tab, text="Sort Order:").grid(row=1, column=0, sticky='w', pady=5)
        sort_order_frame = ttk.Frame(sorting_tab)
        sort_order_frame.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Radiobutton(sort_order_frame, text="Ascending", variable=self.sort_order, value=True).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(sort_order_frame, text="Descending", variable=self.sort_order, value=False).pack(side=tk.LEFT, padx=5)


        # --- Tab 5: Missing Values & Review --- (MOVED FROM 3)
        missing_values_tab = ttk.Frame(self.notebook, padding="15 15 15 15")
        self.notebook.add(missing_values_tab, text="5. Missing Values & Review") # Updated tab number
        missing_values_tab.grid_columnconfigure(0, weight=1)
        missing_values_tab.grid_rowconfigure(1, weight=1)

        ttk.Label(missing_values_tab, text="Rows with Empty Cells (click cell to edit):",
                  font=('Helvetica', 10, 'bold'), foreground=self.fg_color).grid(row=0, column=0, sticky='w', pady=(0,5))

        # Treeview for missing values
        tree_frame = ttk.Frame(missing_values_tab)
        tree_frame.grid(row=1, column=0, sticky='nsew', pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.missing_values_tree = ttk.Treeview(tree_frame, show='headings')
        self.missing_values_tree.grid(row=0, column=0, sticky='nsew')

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.missing_values_tree.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.missing_values_tree.xview)
        hsb.grid(row=1, column=0, sticky='ew')

        self.missing_values_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.missing_values_tree.bind("<Button-1>", self._on_treeview_click) # Bind click for editing


        # --- Process Button (outside tabs) ---
        self.process_button = ttk.Button(main_frame, text="Process File", command=self.process_file)
        self.process_button.pack(pady=20)

        # --- Progress and Status (outside tabs) ---
        self.progress_status_label = ttk.Label(main_frame, text="Ready.", font=('Helvetica', 10, 'bold'), foreground=self.fg_color)
        self.progress_status_label.pack(pady=(0, 5), fill='x')

        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=(0, 10), fill='x')

        self.status_label = ttk.Label(main_frame, text="Select a file and folder to begin.",
                                      font=('Helvetica', 10, 'italic'), wraplength=600)
        self.status_label.pack(pady=(5, 0), fill='x')

        # Try to load initial columns if a file path is already set from config
        if self.input_file_path.get() and os.path.exists(self.input_file_path.get()):
            # Important: now load the actual DF here so it's available for all steps
            self._load_dataframe_and_update_widgets(self.input_file_path.get())


    def _on_mousewheel_date_format(self, event):
        """Allows mousewheel scrolling for the date format canvas."""
        self.date_columns_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_mousewheel_duplicate_check(self, event):
        """Allows mousewheel scrolling for the duplicate check canvas."""
        self.duplicate_check_columns_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_treeview_click(self, event):
        """Handles clicks on the Treeview to enable editing."""
        if self.df is None: return

        item = self.missing_values_tree.identify_row(event.y)
        column_id = self.missing_values_tree.identify_column(event.x)

        if not item or not column_id:
            return

        # Get column name from Treeview's internal column ID
        column_index = int(column_id.replace('#', '')) - 1 # Column ID is #1, #2 etc.
        if column_index < 0 or column_index >= len(self.df.columns):
            return # Should not happen if Treeview is correctly populated

        col_name = self.df.columns[column_index]
        
        # The 'item' variable already holds the iid, which is the original DataFrame index (as a string)
        original_df_index = item 
        
        # Check if the actual value in the DataFrame is NaN
        # We only want to open the editor if the cell is currently empty (NaN)
        if pd.isna(self.df.loc[int(original_df_index), col_name]):
            # Get bounding box of the cell
            x, y, width, height = self.missing_values_tree.bbox(item, column_id)

            # Create an Entry widget over the cell
            entry = ttk.Entry(self.missing_values_tree, style='TEntry') # Use 'TEntry' style
            entry.place(x=x, y=y, width=width, height=height, anchor='nw')
            # Get the actual value from the DataFrame
            current_df_value = self.df.loc[int(original_df_index), col_name]

            # Insert an empty string if the value is missing (NaN, pd.NA, None, NaT), otherwise convert to string
            entry.insert(0, "" if pd.isna(current_df_value) else str(current_df_value))

            entry.focus_set()

            def save_edit(event=None):
                new_value_str = entry.get() # Get the value as a string from the Entry widget

                # Determine the appropriate value to store in the DataFrame
                value_to_store = pd.NA # Default to missing value (pandas' preferred NaN)

                if new_value_str != "": # If the user typed something
                    try:
                        # Try to convert to integer first
                        value_to_store = int(new_value_str)
                    except ValueError:
                        try:
                            # If integer conversion fails, try float
                            value_to_store = float(new_value_str)
                        except ValueError:
                            # If neither int nor float, keep it as a string
                            value_to_store = new_value_str

                # Update the DataFrame using .at for efficient single-cell assignment
                self.df.at[int(original_df_index), col_name] = value_to_store
                
                # Update the Treeview to reflect the change
                # Retrieve all values for the row from the updated DataFrame
                updated_row_values = self.df.loc[int(original_df_index)].values.tolist()
                
                # Convert NaNs (including pd.NA) in the list to empty strings for Treeview display
                display_values = ["" if pd.isna(val) else val for val in updated_row_values]
                
                self.missing_values_tree.item(item, values=display_values)
                
                entry.destroy() # Remove the Entry widget

            entry.bind("<Return>", save_edit) # Save on Enter key
            entry.bind("<FocusOut>", save_edit) # Save when focus is lost (e.g., clicking away)


    def browse_input_file(self):
        initial_dir = os.path.dirname(self.input_file_path.get()) if self.input_file_path.get() else os.path.expanduser('~')
        file_types = [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(
            title="Select Input File",
            initialdir=initial_dir,
            filetypes=file_types
        )
        if file_path:
            self.input_file_path.set(file_path)
            self.update_status(f"Input file selected: {os.path.basename(file_path)}")
            self._save_config()
            # Now, load the entire DataFrame and then update all widgets
            self._load_dataframe_and_update_widgets(file_path)

    def browse_output_folder(self):
        initial_dir = self.output_folder_path.get() if self.output_folder_path.get() else os.path.expanduser('~')
        folder_path = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=initial_dir
        )
        if folder_path:
            self.output_folder_path.set(folder_path)
            self.update_status(f"Output folder selected: {os.path.basename(folder_path)}")
            self._save_config()

    def _load_dataframe_and_update_widgets(self, file_path):
        """Loads the entire DataFrame and then updates all column-dependent widgets."""
        self.update_progress_status("Loading entire file...", 0)
        try:
            if file_path.lower().endswith('.csv'):
                self.df = pd.read_csv(file_path)
            elif file_path.lower().endswith('.xlsx'):
                self.df = pd.read_excel(file_path)
            else:
                self.update_progress_status("Unsupported file type for loading.", 0)
                messagebox.showerror("File Type Error", "Unsupported file type. Please select a .csv or .xlsx file.")
                self.df = None # Clear df if unsupported
                return

            if self.df.empty:
                messagebox.showwarning("Empty File", "The selected file is empty.")
                self.update_progress_status("Empty file loaded.", 0)
                self.df = None
                return

            columns = self.df.columns.tolist()
            
            # --- Update Date Column Checkboxes ---
            for widget in self.date_columns_frame.winfo_children():
                widget.destroy()
            self.date_column_vars.clear()
            
            if not columns:
                 ttk.Label(self.date_columns_frame, text="No columns found in file.", background=self.bg_color).pack(pady=10)
            else:
                col_idx = 0
                for col in columns:
                    var = tk.BooleanVar()
                    if any(keyword in col.lower() for keyword in ["date", "day", "month"]):
                        var.set(True)
                    self.date_column_vars[col] = var
                    chk = ttk.Checkbutton(self.date_columns_frame, text=col, variable=var)
                    chk.grid(row=col_idx // 3, column=col_idx % 3, sticky='w', padx=2, pady=2)
                    col_idx += 1
                self.date_columns_canvas.update_idletasks()
                self.date_columns_canvas.config(scrollregion=self.date_columns_canvas.bbox("all"))

            # --- Update Duplicate Check Column Checkboxes ---
            for widget in self.duplicate_check_columns_frame.winfo_children():
                widget.destroy()
            self.duplicate_check_column_vars.clear()

            if not columns:
                 ttk.Label(self.duplicate_check_columns_frame, text="No columns found in file.", background=self.bg_color).pack(pady=10)
            else:
                col_idx = 0
                for col in columns:
                    var = tk.BooleanVar(value=True) # Default to all columns selected for duplicate check
                    self.duplicate_check_column_vars[col] = var
                    chk = ttk.Checkbutton(self.duplicate_check_columns_frame, text=col, variable=var)
                    chk.grid(row=col_idx // 3, column=col_idx % 3, sticky='w', padx=2, pady=2)
                    col_idx += 1
                self.duplicate_check_columns_canvas.update_idletasks()
                self.duplicate_check_columns_canvas.config(scrollregion=self.duplicate_check_columns_canvas.bbox("all"))

            # --- Update Missing Values Treeview ---
            self._populate_missing_values_treeview()

            # --- Update Sort Column Combobox ---
            self.sort_column_combobox['values'] = columns
            if columns:
                if self.sort_column.get() not in columns:
                    self.sort_column.set(columns[0])
            else:
                self.sort_column.set("")

            self.update_progress_status("File loaded and options updated.", 10)

        except Exception as e:
            messagebox.showerror("Error Loading File", f"Could not load or read file: {e}")
            self.update_progress_status("Error loading file.", 0)
            self.df = None # Clear df on error
            # Clear all dependent widgets
            for widget in self.date_columns_frame.winfo_children(): widget.destroy()
            self.date_column_vars.clear()
            for widget in self.duplicate_check_columns_frame.winfo_children(): widget.destroy()
            self.duplicate_check_column_vars.clear()
            self._clear_missing_values_treeview()
            self.sort_column_combobox['values'] = []
            self.sort_column.set("")

    def _populate_missing_values_treeview(self):
        """Populates the Treeview with rows containing NaN values."""
        self._clear_missing_values_treeview() # Clear previous data

        if self.df is None or self.df.empty:
            return

        # Identify rows with any NaN values
        missing_rows_df = self.df[self.df.isnull().any(axis=1)]

        if missing_rows_df.empty:
            self.missing_values_tree.heading("#0", text="") # Clear default heading
            self.missing_values_tree["columns"] = () # Clear column definitions
            ttk.Label(self.missing_values_tree.master, text="No empty cells found in file.", background=self.bg_color, foreground=self.fg_color).pack(pady=10)
            return

        # Define Treeview columns
        columns = self.df.columns.tolist()
        self.missing_values_tree["columns"] = columns
        self.missing_values_tree.column("#0", width=0, stretch=tk.NO) # Hide default first column
        
        for col in columns:
            self.missing_values_tree.heading(col, text=col)
            self.missing_values_tree.column(col, width=100, anchor='w') # Default width

        # Insert data
        for original_idx, row in missing_rows_df.iterrows():
            # Convert NaNs to empty strings for Treeview display
            display_values = ["" if pd.isna(val) else val for val in row.values.tolist()]
            self.missing_values_tree.insert("", "end", iid=str(original_idx), values=display_values) # iid is original df index

    def _clear_missing_values_treeview(self):
        """Clears all data from the missing values Treeview."""
        if self.missing_values_tree:
            for item in self.missing_values_tree.get_children():
                self.missing_values_tree.delete(item)
            self.missing_values_tree["columns"] = () # Clear column definitions
            self.missing_values_tree.heading("#0", text="") # Clear placeholder text

    def process_file(self):
        # Use self.df directly which has been loaded and potentially edited
        if self.df is None:
            messagebox.showwarning("No File Loaded", "Please load an input file first.")
            self.update_progress_status("Error: No file loaded.", 0)
            return

        output_folder = self.output_folder_path.get()
        selected_strftime_format = self._date_format_map.get(self.selected_date_format_display.get())
        sort_col = self.sort_column.get()
        sort_asc = self.sort_order.get()

        # --- Validation ---
        if not output_folder:
            messagebox.showwarning("Output Error", "Please select an output folder.")
            self.update_progress_status("Error: No output folder selected.", 0)
            return
        if not selected_strftime_format:
            messagebox.showwarning("Date Format Error", "Please select a valid date format from the dropdown.")
            self.update_progress_status("Error: No date format selected.", 0)
            return
        if self.do_remove_duplicates.get():
            selected_duplicate_columns = [col_name for col_name, var in self.duplicate_check_column_vars.items() if var.get()]
            if not selected_duplicate_columns:
                messagebox.showwarning("Duplicate Check Error", "If 'Remove Duplicate Rows' is checked, you must select at least one column for duplicate checking.")
                self.update_progress_status("Error: No columns selected for duplicate check.", 0)
                return

        self.update_progress_status("Starting processing...", 0)
        self.root.update_idletasks()

        try:
            # Create a copy of the DataFrame to apply cleaning steps, preserving self.df for re-runs
            processed_df = self.df.copy()
            original_rows = len(processed_df)
            current_progress = 20

            # --- Step 1: Trim Whitespace (Conditional) ---
            if self.do_trim_whitespace.get():
                self.update_progress_status("Trimming whitespace...", current_progress)
                for col in processed_df.select_dtypes(include=['object']).columns:
                    # Apply .str.strip() directly. This handles NaNs correctly by leaving them as NaN.
                    processed_df[col] = processed_df[col].str.strip() # <--- CORRECTED
            current_progress += 10


            # --- Step 2: Capitalize String Columns (Conditional) ---
            if self.do_capitalize_strings.get():
                self.update_progress_status("Capitalizing text fields...", current_progress)
                for col in processed_df.select_dtypes(include=['object']).columns:
                    # Apply .str.title() directly. This handles NaNs correctly by leaving them as NaN.
                    processed_df[col] = processed_df[col].str.title() # <--- CORRECTED
            current_progress += 10


            # --- Step 3: Date Column Processing ---
            processed_date_cols = []
            self.update_progress_status("Formatting selected date columns...", current_progress)
            for col_name, var in self.date_column_vars.items():
                if var.get() and col_name in processed_df.columns:
                    processed_date_cols.append(col_name)
                    try:
                        processed_df[col_name] = pd.to_datetime(processed_df[col_name], errors='coerce')
                        processed_df[col_name] = processed_df[col_name].dt.strftime(selected_strftime_format).fillna('')
                    except Exception as e:
                        messagebox.showwarning("Date Formatting Warning",
                                               f"Could not apply format to column '{col_name}'. "
                                               f"Error: {e}\nInvalid dates will be cleared.")
                        processed_df[col_name] = processed_df[col_name].fillna('')
            current_progress += 20

            # --- Step 4: Duplicate Detection and Removal (Conditional) ---
            num_removed_duplicates = 0
            if self.do_remove_duplicates.get():
                self.update_progress_status("Detecting and removing duplicates...", current_progress)
                subset_cols = [col_name for col_name, var in self.duplicate_check_column_vars.items() if var.get()]
                valid_subset_cols = [col for col in subset_cols if col in processed_df.columns]
                
                if not valid_subset_cols:
                    messagebox.showwarning("Duplicate Check Warning", "No valid columns selected for duplicate checking. Skipping duplicate removal.")
                else:
                    processed_df = processed_df.drop_duplicates(subset=valid_subset_cols, keep='first')
                    num_removed_duplicates = original_rows - len(processed_df)
            
            cleaned_rows = len(processed_df)
            current_progress += 20

            # --- Step 5: Sort Data ---
            if sort_col and sort_col in processed_df.columns:
                self.update_progress_status(f"Sorting data by '{sort_col}'...", current_progress)
                processed_df = processed_df.sort_values(by=sort_col, ascending=sort_asc)
            current_progress += 5

            # --- Step 6: Save Cleaned DataFrame ---
            self.update_progress_status("Saving cleaned file...", current_progress)
            original_filename_with_ext = os.path.basename(self.input_file_path.get())
            original_name_without_ext, _ = os.path.splitext(original_filename_with_ext)
            cleaned_filename = f"cleaned.{original_name_without_ext}.xlsx"
            output_file_path = os.path.join(output_folder, cleaned_filename)
            processed_df.to_excel(output_file_path, index=False)
            current_progress = 100

            # --- Final Status and Message ---
            self.update_progress_status("Processing complete!", current_progress)

            summary_message = (
                f"Successfully processed:\n"
                f"Original rows: {original_rows}\n"
            )
            if processed_date_cols:
                summary_message += f"Processed date columns: {', '.join(processed_date_cols)} (formatted to '{self.selected_date_format_display.get()}')\n"
            else:
                summary_message += "No 'date' columns selected or processed.\n"

            if self.do_remove_duplicates.get():
                if num_removed_duplicates > 0:
                    summary_message += (
                        f"Duplicates removed: {num_removed_duplicates}\n"
                        f"Final rows: {cleaned_rows}\n"
                        f"Cleaned file saved to:\n{output_file_path}"
                    )
                else:
                    summary_message += (
                        f"No duplicates found or removed (based on selected columns).\n"
                        f"Final rows: {cleaned_rows}\n"
                        f"Cleaned file saved to:\n{output_file_path}"
                    )
            else:
                 summary_message += (
                    f"Duplicate removal was skipped.\n"
                    f"Final rows: {cleaned_rows}\n"
                    f"Cleaned file saved to:\n{output_file_path}"
                )

            if sort_col:
                summary_message += f"\nData sorted by '{sort_col}' ({'Ascending' if sort_asc else 'Descending'})."

            messagebox.showinfo("Processing Complete", summary_message)

        except FileNotFoundError:
            messagebox.showerror("Error", "The specified file was not found.")
            self.update_progress_status("Error: File not found.", 0)
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", "The selected file is empty or corrupted.")
            self.update_progress_status("Error: Empty or corrupted file.", 0)
        except Exception as e:
            messagebox.showerror("An Error Occurred", f"An unexpected error occurred: {e}")
            self.update_progress_status(f"Error: {e}", 0)
        finally:
            self.progress_bar['value'] = 0
            self.root.update_idletasks()


    def update_progress_status(self, message, value):
        self.progress_status_label.config(text=message)
        self.progress_bar['value'] = value
        self.root.update_idletasks()

    def update_status(self, message):
        self.status_label.config(text=message)

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = CleanItApp(root)
    root.mainloop()
