**CleanIt: Intelligent Data Cleaning & Formatting Tool**
**Project Documentation**

Version: 1.0

Date: July 15th, 2025

**1. Introduction**

CleanIt is a desktop application designed to simplify and automate common data preparation and cleaning tasks for tabular data. It provides a user-friendly graphical interface built with Tkinter that allows users to process messy CSV and Excel files, enhancing data quality and consistency before further analysis or integration.

The primary purpose of CleanIt is to address the common challenges associated with inconsistent data formats, duplicate entries, missing information, and improper text capitalization, offering a streamlined solution to transform raw data into a clean, standardized, and usable format.

**2. Features**

* Intuitive GUI: A clear and organized graphical user interface with a tabbed layout for easy navigation between different cleaning functionalities.
* Flexible File Input: Supports reading and processing data from both Comma Separated Values (.csv) and Microsoft Excel (.xlsx) file formats.
* Configurable Cleaning Operations:
  * Whitespace Trimming: Automatically removes leading and trailing whitespace characters from all text-based fields, ensuring clean data entries.
  * Text Field Capitalization: Converts text fields to Title Case (e.g., "john doe" becomes "John Doe"), standardizing textual data.
  * Duplicate Row Removal: Identifies and removes exact duplicate rows. Users have precise control to define which specific columns should be considered when detecting duplicates, allowing for flexible management (e.g., ignoring unique identifiers like a Transaction ID).
* Advanced Date Formatting:
  * Offers a comprehensive selection of user-friendly output date formats (e.g., "MM/DD/YYYY", "Weekday, Month DD, YYYY") for consistent date representation.
  * Automatically suggests columns likely containing date information (based on common keywords like "Date", "Day", or "Month") for user review and selection.
  * Handles various input date formats and replaces unparseable or invalid date entries with blanks.
* Interactive Missing Values Review & Editing:
  * A dedicated tab displays all rows that contain one or more empty (blank) cells in an interactive table (Treeview).
  * Users can directly click on cells within this table to perform in-place editing, enabling manual correction or filling of missing data.
  * Ensures that missing values are correctly represented as truly blank cells in the final Excel output.
* Data Sorting: Provides the ability to sort the cleaned data by any user-selected column in either ascending or descending order.
* Enhanced User Experience:
  * Remembers the last-used input file path and output folder, minimizing repetitive navigation for recurring tasks.
  * Features a visual progress bar and detailed status updates during data processing, providing clear feedback on the operation's progress.
  * Launches in a maximized window state for an optimized workspace.
  * Robust Error Handling: Incorporates comprehensive error management to gracefully handle file access issues, data conversion problems, and other unexpected runtime events.

**3. How to Use**

Follow these steps to clean your data using the CleanIt application:

**Launch the Application:**

Navigate to the directory containing cleanit_app.py.
Execute the script using Python.
python cleanit_app.py

**1. File & Output Tab:**

* Input File: Click the "Browse..." button next to "Input File:" to select the CSV or Excel file you wish to clean.
* Output Folder: Click the "Browse..." button next to "Output Folder:" to choose the destination directory where the cleaned file will be saved.

**2. Cleaning Options Tab:**

* Global Cleaning: Check or uncheck the boxes to enable/disable "Trim Leading/Trailing Whitespace" and "Capitalize Text Fields (Title Case)".
* Duplicate Removal:
  * Check "Remove Duplicate Rows" to enable this feature.
  * Below this, carefully select the columns that CleanIt should use to identify duplicate rows. Only rows where all selected columns match will be considered duplicates. For example, if "TransactionID" is unique but all other columns are the same, you would uncheck "TransactionID" to mark such rows as duplicates.

**3. Date Formatting Tab:**

* Output Format: Select your preferred output date format from the "Choose Output Date Format:" dropdown list. Examples are provided for clarity.
* Column Selection: Review the automatically suggested columns (those with "Date", "Day", or "Month" in their names) and check/uncheck them based on your requirements. Only checked columns will be formatted.

**4. Sorting Tab:**

* Sort Column: Select the column by which you want to sort the entire dataset from the "Sort Data By Column:" dropdown.
* Sort Order: Choose "Ascending" or "Descending" for the sort order.

**5. Missing Values & Review Tab:**

* Once a file is loaded, this tab will automatically populate with all rows that contain one or more blank cells.
* Interactive Editing: To fill a blank cell, click on it. An editable text box will appear. Type your desired value and press Enter or click anywhere else to save the change to the underlying data. If you leave the box blank and click away, the cell will remain empty.

* Process File:
  * After configuring all desired options, click the prominent "Process File" button at the bottom of the application window.
  * A progress bar and status messages will provide real-time updates.
  * Upon completion, a summary message box will appear, and the cleaned file (named cleaned.[original_filename].xlsx) will be saved in your selected output folder.

**4. Installation**

* Prerequisites: Ensure Python 3.8 is installed on your system.
* Install Dependencies: Open a terminal or command prompt in the directory where cleanit_app.py is located and run the following command:
  * pip install pandas openpyxl

**5. Technical Stack**

* Python 3.8: The core programming language.
* Tkinter (with ttk): Python's standard GUI toolkit for building the interactive user interface. ttk provides themed widgets for a modern look.
* Pandas: A powerful and widely used data manipulation and analysis library, essential for reading/writing files, performing data cleaning operations, and structured data handling.
* os module: A standard Python library module used for interacting with the operating system, primarily for file path manipulation and directory operations.
* json module: A standard Python library module employed for saving and loading user preferences (such as last-used file paths) to a configuration file, enhancing user convenience.

**6. Testing with Provided Data**

To thoroughly test all functionalities of the CleanIt application, please use the test_data.xlsx file provided alongside this documentation.

* This file is specifically designed to demonstrate:

  * Duplicate Rows: Including scenarios where a unique identifier (like TransactionID) needs to be excluded from the duplicate check.
  * Mixed Date Formats: Various date representations that CleanIt should standardize.
  * Whitespace: Cells with leading or trailing spaces.
  * Mixed Case Text: Text fields with inconsistent capitalization.
  * Missing Values: Blank cells that can be interactively filled in the "Missing Values & Review" tab.

**Suggested Test Scenarios with test_data.xlsx:**

* Load test_data.xlsx.
* Navigate to the "Cleaning Options" tab and uncheck "TransactionID" in the duplicate check columns. Then, process the file to observe duplicate "Alice" and "Bob" rows being removed.
* Experiment with different output formats in the "Date Formatting" tab.
* Access the "Missing Values & Review" tab and manually fill in some blank cells (e.g., for David's Quantity or the empty "Notes" cell).
* Test the sorting functionality by selecting various columns and sort orders.
