import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import sys # For checking platform
from datetime import datetime # To generate example dates

class CleanItApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CleanIt")

        # --- Fullscreen/Maximized Window ---
        try:
            self.root.state('zoomed')
        except tk.TclError:
            pass # Fallback for systems where 'zoomed' state is not supported

        self.root.resizable(True, True) # Allow resizing after initial maximize

        # --- Colors ---
        self.bg_color = "#E6F7FF"  # Light Blue (background)
        self.fg_color = "#000080"  # Dark Blue (for text)

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


        # --- Variables to store file paths and date format ---
        self.input_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()

        # --- Reference date for examples ---
        # Using a fixed date like 2023-01-15 13:30:45 (Sunday) for consistent examples
        example_date = datetime(2023, 1, 15, 13, 30, 45)

        # --- Define date format options (display_string, strftime_code) ---
        # The order here determines the order in the dropdown
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

        # Generate the list of display strings for the combobox
        self.date_format_display_strings = [item[0] for item in self._date_format_options]

        # Create a mapping from display string to strftime code for easy lookup
        self._date_format_map = {item[0]: item[1] for item in self._date_format_options}

        self.selected_date_format_display = tk.StringVar(value=self.date_format_display_strings[0]) # Default to the first in the list

        self.create_widgets()

    def create_widgets(self):
        # --- Main Frame ---
        main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        main_frame.pack(expand=True, fill='both')

        # --- Title ---
        title_label = ttk.Label(main_frame, text="CleanIt: Detect & Clean Duplicates and Format Dates",
                                font=('Helvetica', 16, 'bold'), foreground=self.fg_color)
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky='ew')

        # --- Input File Selection ---
        ttk.Label(main_frame, text="1. Select Input File:").grid(row=1, column=0, sticky='w', pady=5)
        self.input_entry = ttk.Entry(main_frame, textvariable=self.input_file_path, state='readonly')
        self.input_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(main_frame, text="Browse...", command=self.browse_input_file).grid(row=1, column=2, padx=5, pady=5)

        # --- Output Folder Selection ---
        ttk.Label(main_frame, text="2. Select Output Folder:").grid(row=2, column=0, sticky='w', pady=5)
        self.output_entry = ttk.Entry(main_frame, textvariable=self.output_folder_path, state='readonly')
        self.output_entry.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(main_frame, text="Browse...", command=self.browse_output_folder).grid(row=2, column=2, padx=5, pady=5)

        # --- Date Format Selection ---
        ttk.Label(main_frame, text="3. Choose Output Date Format:").grid(row=3, column=0, sticky='w', pady=5)
        self.date_format_combobox = ttk.Combobox(main_frame, textvariable=self.selected_date_format_display,
                                                 values=self.date_format_display_strings, state='readonly')
        self.date_format_combobox.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        # Set default value explicitly
        if self.date_format_display_strings:
            self.selected_date_format_display.set(self.date_format_display_strings[0])


        # --- Process Button ---
        self.process_button = ttk.Button(main_frame, text="4. Process File", command=self.process_file)
        self.process_button.grid(row=4, column=0, columnspan=3, pady=20)

        # --- Status Label ---
        self.status_label = ttk.Label(main_frame, text="Ready. Select a file and folder to begin.",
                                      font=('Helvetica', 10, 'italic'), wraplength=600)
        self.status_label.grid(row=5, column=0, columnspan=3, pady=10, sticky='ew')

        # --- Configure column weights for resizing ---
        main_frame.grid_columnconfigure(1, weight=1) # Allow the entry and combobox fields to expand
        for i in range(6): # Allow rows to expand
            main_frame.grid_rowconfigure(i, weight=1)


    def browse_input_file(self):
        file_types = [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(
            title="Select Input File",
            filetypes=file_types
        )
        if file_path:
            self.input_file_path.set(file_path)
            self.update_status(f"Input file selected: {os.path.basename(file_path)}")

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select Output Folder"
        )
        if folder_path:
            self.output_folder_path.set(folder_path)
            self.update_status(f"Output folder selected: {os.path.basename(folder_path)}")

    def process_file(self):
        input_file = self.input_file_path.get()
        output_folder = self.output_folder_path.get()
        
        # Get the selected display string and map it to the actual strftime code
        selected_display_string = self.selected_date_format_display.get()
        selected_strftime_format = self._date_format_map.get(selected_display_string)

        if not input_file:
            messagebox.showwarning("Input Error", "Please select an input file.")
            self.update_status("Error: No input file selected.")
            return

        if not output_folder:
            messagebox.showwarning("Output Error", "Please select an output folder.")
            self.update_status("Error: No output folder selected.")
            return

        if not selected_strftime_format:
            messagebox.showwarning("Date Format Error", "Please select a valid date format from the dropdown.")
            self.update_status("Error: No date format selected.")
            return

        self.update_status("Processing file... Please wait.")
        self.root.update_idletasks() # Update GUI to show "Processing..." immediately

        try:
            # Determine file type and read
            if input_file.lower().endswith('.csv'):
                df = pd.read_csv(input_file)
            elif input_file.lower().endswith('.xlsx'):
                df = pd.read_excel(input_file)
            else:
                messagebox.showerror("File Type Error", "Unsupported file type. Please select a .csv or .xlsx file.")
                self.update_status("Error: Unsupported file type.")
                return

            original_rows = len(df)
            processed_date_cols = []

            # --- Date Column Processing ---
            for col in df.columns:
                if "date" in str(col).lower(): # Check if 'date' is in the column name (case-insensitive)
                    processed_date_cols.append(col)
                    self.update_status(f"Attempting to format date column: '{col}'...")
                    self.root.update_idletasks()
                    try:
                        # Convert to datetime objects, coercing errors to NaT (Not a Time)
                        df[col] = pd.to_datetime(df[col], errors='coerce')
                        # Format back to string, NaT values will become blank strings
                        df[col] = df[col].dt.strftime(selected_strftime_format).fillna('')
                    except Exception as e:
                        # If formatting fails for some reason (e.g., malformed format string, though unlikely with pre-defined)
                        messagebox.showwarning("Date Formatting Warning",
                                               f"Could not apply format '{selected_display_string}' to column '{col}'. "
                                               f"Error: {e}\nInvalid dates will be cleared.")
                        self.update_status(f"Warning: Date formatting failed for '{col}'. Error: {e}")
                        df[col] = df[col].fillna('') # Ensure NaT becomes empty string

            # --- Duplicate Detection and Removal ---
            
            # Detect duplicates (keeping all instances for counting)
            duplicate_rows_mask = df.duplicated(keep=False)
            num_detected_duplicates = duplicate_rows_mask.sum()

            # Remove duplicates, keeping the first occurrence
            cleaned_df = df.drop_duplicates(keep='first')
            cleaned_rows = len(cleaned_df)
            num_removed_duplicates = original_rows - cleaned_rows

            # Construct output filename
            original_filename_with_ext = os.path.basename(input_file)
            original_name_without_ext, _ = os.path.splitext(original_filename_with_ext)
            
            cleaned_filename = f"cleaned.{original_name_without_ext}.xlsx"
            output_file_path = os.path.join(output_folder, cleaned_filename)

            # Save the cleaned DataFrame
            cleaned_df.to_excel(output_file_path, index=False)

            # --- Success Message ---
            summary_message = (
                f"Successfully processed:\n"
                f"Original rows: {original_rows}\n"
            )
            if processed_date_cols:
                summary_message += f"Processed date columns: {', '.join(processed_date_cols)} (formatted to '{selected_display_string}')\n"
            else:
                summary_message += "No 'date' columns found or processed.\n"

            if num_removed_duplicates > 0:
                summary_message += (
                    f"Duplicates removed: {num_removed_duplicates}\n"
                    f"Cleaned rows: {cleaned_rows}\n"
                    f"Cleaned file saved to:\n{output_file_path}"
                )
                self.update_status(f"Done! {num_removed_duplicates} duplicates removed, dates formatted. File saved to: {output_file_path}")
            else:
                summary_message += (
                    f"No duplicates found.\n"
                    f"Cleaned file (identical to original, with date formatting) saved to:\n{output_file_path}"
                )
                self.update_status(f"Done! No duplicates found, dates formatted. File saved to: {output_file_path}")

            messagebox.showinfo("Processing Complete", summary_message)

        except FileNotFoundError:
            messagebox.showerror("Error", "The specified file was not found.")
            self.update_status("Error: File not found.")
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", "The selected file is empty or corrupted.")
            self.update_status("Error: Empty or corrupted file.")
        except Exception as e:
            messagebox.showerror("An Error Occurred", f"An unexpected error occurred: {e}")
            self.update_status(f"Error: {e}")

    def update_status(self, message):
        self.status_label.config(text=message)

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = CleanItApp(root)
    root.mainloop()
