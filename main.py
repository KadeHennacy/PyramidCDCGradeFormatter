# Kade Hennacy 9/4/2024

# To best understand this script, read these comments in order starting at 1. 2 is at the bottom. Alt + Z enables word wrap in VSCode so you can read more easily

# 1: These are the libraries we use. os (operating system), re(regular expressions), and tkinter(TK interface) are included with the standard python installation so aren't included in requirements.txt
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button, Frame, IntVar, Checkbutton
from tkinter.ttk import Combobox, Spinbox
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font


# 21: This function is called when load_button is pressed
def load_file():
    # 22: declare file_path as global so we can access it anywhere in this script.
    global file_path
    # 23: Get the format setting user has selected from the dropdown combo
    format_setting = format_combo.get()
    
    # 24: If the setting is Gmetrix, only accept CSV, else except either format
    if format_setting == "Gmetrix Raw Data":
        file_type = [("CSV files", "*.csv")]
    else:  # 25: General Formatting
        file_type = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]

    # 26: We create a popup filepicker and save the path they select as a global variable so we can access it anywhere in this code
    file_path = filedialog.askopenfilename(filetypes=file_type)
    if file_path:
        filename = os.path.basename(file_path)
        # 27: Display the name of the loaded file on label.
        file_label.config(text=f"Loaded file: {filename}", font=('Helvetica', 10, 'bold'))

# 28: This function is called when save_button is pressed
def save_file():
    # 29: This determines whether to read as CSV or Excel and how to process it. Step into this for comment #30
    process_file()
    if 'df' not in globals():
        messagebox.showerror("Error", "No processed data available. Please load and process a CSV file first.")
        return

    # 41: This creates a popup to ask the user where to save the file and what to name it. 
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_file_path:
        return

    # 42: Depending on the file type, we use openpyxl to create a workbook and worksheet. Then we populate its cells from the Pandas dataframe.
    if file_path.endswith('.xlsx'):
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active


    for row_idx, row in enumerate(df.itertuples(index=False, name=None), 1):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # 43: This applies the general formatting like word wrap, centering text, and resizing columns. Step into this for comment  #44
    general_formatting(ws)

    # 50: Save the file and display success message
    wb.save(output_file_path)
    messagebox.showinfo("Success", f"Data processed and saved to {output_file_path}")

def process_file():
    # 30: We declare df (Dataframe) as a global variable so we can access it anywhere in this code. We use Pandas to create this from either a CSV or XLXS file.
    global df, file_path
    format_setting = format_combo.get()

    try:
        if file_path.endswith('.csv'):
            # 31: Pandas can only read CSVs with an equal amount of commas in each row. Gmetrix export is missing these because it doesn't include commas for empty cells. Step into sanitize_csv() for #32
            sanitize_csv()
            df = pd.read_csv(file_path, header=None)
        else:
            df = pd.read_excel(file_path, header=None)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process the file\n{e}")
        return
    # 33: We run general formatting regardless, but Gmetrix requires sorting. Step into process_gmetrix for comment #34
    if format_setting == "Gmetrix Raw Data":
        process_gmetrix()

def process_gmetrix():
    global df
    # 34: First, check if the dataframe is loaded, otherwise show an error.
    if df is None:
        messagebox.showerror("Error", "No file loaded. Please load a CSV file first.")
        return

    # 35: Identify columns that contain "Minutes Spent" or "Score" and mark them for removal.
    columns_to_remove = []
    for col in df.columns:
        if df[col].apply(lambda x: str(x).strip().lower() in ["minutes spent", "score"]).any():
            columns_to_remove.append(col)

    # 36: Drop the identified columns from the dataframe to clean the dataset.
    df.drop(columns_to_remove, axis=1, inplace=True)

    # 37: Check the sort order preference from the combo box and prepare to sort data if required.
    sort_order = sort_order_combo.get()
    if sort_order != "Unsorted":
        ascending_order = sort_order == "Ascending"
        post_assessment_cols = []

        # 38: Locate columns related to post-assessment to sort the data based on scores.
        for col in df.columns:
            if df[col].apply(lambda x: "Post-Assessment" in str(x)).any():
                post_assessment_cols.append(col)

        # 39: For each post-assessment column, find the rows with 'Test Score' and determine the range to sort.
        for post_assessment_col in post_assessment_cols:
            test_score_rows = df[df[post_assessment_col] == 'Test Score'].index

            for test_score_row in test_score_rows:
                next_blank_row = df[df.index > test_score_row][post_assessment_col].first_valid_index()
                end_row = df[df.index >= next_blank_row][post_assessment_col].isna().idxmax() if pd.notna(next_blank_row) else len(df)

                # 40: Extract score data, convert percentages to numeric, sort, and reintegrate into the dataframe. Return to save_file() for comment # 41
                scores_data = df.iloc[test_score_row + 1:end_row].copy()
                scores_data[post_assessment_col] = scores_data[post_assessment_col].str.rstrip('%').apply(pd.to_numeric, errors='coerce')
                scores_data.dropna(subset=[post_assessment_col], inplace=True)
                sorted_data = scores_data.sort_values(by=post_assessment_col, ascending=ascending_order)
                sorted_data[post_assessment_col] = sorted_data[post_assessment_col].apply(lambda x: f"{x}%")
                df.iloc[test_score_row + 1:end_row] = sorted_data


# 32: We count the number of commas in the line with the most, and add the difference to any line lacking commas to make them all equal. Return to process_file() for #33
def sanitize_csv():
    global file_path

    with open(file_path, 'r') as file:
        lines = file.readlines()

    max_commas = max(line.count(',') for line in lines)

    adjusted_lines = []
    for line in lines:
        current_commas = line.count(',')
        if current_commas < max_commas:
            line = line.strip('\n') + ',' * (max_commas - current_commas) + '\n'
        adjusted_lines.append(line)

    with open(file_path, 'w') as file:
        file.writelines(adjusted_lines)


# 44: We pass in the worksheet as an argument and iterate through each column
def general_formatting(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter

        for cell in col:
            # 45: Apply word wrap
            if word_wrap_var.get() == 1:
                cell.alignment = Alignment(wrap_text=True)
            
            # 46: Apply text centering
            if center_text_var.get() == 1:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True) if cell.alignment.wrap_text else Alignment(horizontal='center', vertical='center')
            
            # 47: Calculate max length of the column content
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        # 48: Autosize column based on content if autosize is active
        if autosize_col_var.get() == 1:
            ws.column_dimensions[column].width = max_length + 2  
        # 49: Manually size column to user defined width if manual resize is active. Return to save_file() for #50
        elif resize_col_var.get() == 1:
            ws.column_dimensions[column].width = column_width_var.get()

# 3: This is where we initialize the main window of the app "root".
# All elements will be attached to this root
root = tk.Tk()
root.title("Spreadsheet Formatter")

# 4: This defines the default size of the window
root.geometry("900x400")

# 5: Frame(root) creates a new section and pack attaches it to the main window. The order they're packed determines the order they appear on the page In this case frame_top is the top section, frame_sort is the middle, and frame_bottom is the bottom pady determines the space above and below the section. So pady=(10,0) means 10px above, 0 below
frame_top = Frame(root)
frame_top.pack(pady=10)

frame_sort = Frame(root)
frame_sort.pack(pady=(10, 0))

frame_bottom = Frame(root)
frame_bottom.pack(pady=0)

# 6: We create a label in the bottom of the window to show the name of the selected file. It's not visible until loadFile is called which sets the value with file_label.config(text=f"Loaded file: {filename}"...
file_label = Label(frame_bottom, wraplength=600, justify="left")
file_label.pack(pady=10) 

# 7: We add general instructions. This displays underneath file_label because it's packed afterward
instruction_text = "This program loads a spreadsheet and formats it according to the format setting."
label_instructions = Label(frame_bottom, text=instruction_text, wraplength=600, justify="left")
label_instructions.pack(pady=(30, 15))


# 8: We add specific instructions based on the format setting. The message is set in update_instructions() 
additional_instruction_label = Label(frame_bottom, wraplength=600, justify="left")
additional_instruction_label.pack(pady=10)

format_setting_label = Label(frame_top, text="Format Setting")
format_setting_label.pack(side=tk.LEFT, padx=10)

# 9: This defines the 2 format settings the app supports. The value here determines what update_instructions() will set additional_instruction_label to This determines if "Sort Order" setting is supported. This determines if XLSX files may be loaded
format_combo = Combobox(frame_top, state="readonly", width=20)
format_combo['values'] = ("Gmetrix Raw Data", "General Formatting")
format_combo.current(1)
format_combo.pack(side=tk.LEFT, padx=10)

# 10: This defines the sort order setting. This is only visible if format_combo.get() == "Gmetrix Raw Data" in process_gmetrix() it checks the value of this to determine how to sort the rows
sort_order_label = Label(frame_sort, text="Sort Order")
sort_order_combo = Combobox(frame_sort, state="readonly", width=15)
sort_order_combo['values'] = ("Ascending", "Descending", "Unsorted")
sort_order_combo.current(1)

# 11: This calls load_file when it's clicked, which opens a file-picker and stores the path to the file the user picks as a global variable
load_button = Button(frame_top, text="Load Input File", command=load_file)
load_button.pack(side=tk.LEFT, padx=10)

# 12: This calls save_file when it's clicked, which processes the file according to the format settings, opens a file-picker, and saves the result at the specified location
save_button = Button(frame_top, text="Save Formatted File", command=save_file)
save_button.pack(side=tk.LEFT, padx=10)

# 13: This function is called whenever the format setting is changed. It checks which setting was selected and updates the instruction label to display relevant instructions
def update_instruction(event=None):
    if format_combo.get() == "Gmetrix Raw Data":
        additional_instruction_label.config(text="Gmetrix Raw Data - This setting takes a CSV Gmetrix student progress report as input. It removes the \"Minutes Spent\" and \"Score\" columns, sizes columns to 18, and centers and wraps all text, and outputs a XLSX file. To sort the rows by post-assessment score, select an option from the Sort Order dropdown. Descending places the highest scores at the top of the sheet")
        sort_order_label.grid(row=0, column=0, padx=(10, 2), sticky='e')
        sort_order_combo.grid(row=0, column=1, padx=(2, 10), sticky='w')
    elif format_combo.get() == "General Formatting":
        additional_instruction_label.config(text="General Formatting - This setting takes any CSV or XLSX file and sizes columns to 18 and centers and wraps text")
        sort_order_label.grid_remove()
        sort_order_combo.grid_remove()

# 14: If the user selects the "Resize Columns" checkbutton we add a spinbox for the user to select how many points wide they want the columns. We also need to disable "Autosize" columns because the user can only choose one or the other. If resize_col_var.get() == 1 that means  "Resize Columns" is selected, so we show the spinbox for column with, and we show a label for it, and we deselect the "Autosize" checkbutton by setting it to 0. Otherwise, we hide the column size spinbox and points label.
def handle_resize_checkbutton():
    if resize_col_var.get() == 1:
        autosize_col_var.set(0)
        column_width_spin.grid()
        px_label.grid()
    else:
        column_width_spin.grid_remove()
        px_label.grid_remove()

# 15 The block above will disable the Autosize option if you select the Resize option, and this does the reverse. If Resize is selected, and the user selects Autosize, this deselects the resize option and hides the column width spinbox and label. This seems redundant, because above has an 'else' statement that removes these if resize isn't selected, but handle_resize_checkbutton() is only linked to the resize checkbutton. We need another function to link to the autosize checkbutton.
def handle_autosize_checkbutton():
    if autosize_col_var.get() == 1:
        resize_col_var.set(0)
        column_width_spin.grid_remove()
        px_label.grid_remove()


# 16 These variables represent default values for the UI inputs. 1 = Checked for checkbuttons
word_wrap_var = IntVar(value=1)
center_text_var = IntVar(value=1)
autosize_col_var = IntVar(value=1)
resize_col_var = IntVar(value=0)
column_width_var = IntVar(value=18)

# 17 Initialize UI inputs. Attach them to frame_sort (middle). Link any functions that need to be ran on-change with "command" argument
word_wrap_check = Checkbutton(frame_sort, text="Word Wrap", variable=word_wrap_var)
center_text_check = Checkbutton(frame_sort, text="Center Text", variable=center_text_var)
resize_col_check = Checkbutton(frame_sort, text="Resize Columns", variable=resize_col_var, command=handle_resize_checkbutton)
autosize_col_check = Checkbutton(frame_sort, text="Autosize Columns", variable=autosize_col_var, command=handle_autosize_checkbutton)
column_width_spin = Spinbox(frame_sort, from_=10, to=50, textvariable=column_width_var, width=5)
px_label = Label(frame_sort, text="Points")

# 18 Set position of the elments with .grid() to all be on a row. Set padding so they look evenly spaced. .grid() is similar to .pack() in the way it attaches the element to the frame and makes it visible, but .grid() uses rows and columns to give more control over positioning.
word_wrap_check.grid(row=0, column=2, padx=(10, 2), sticky='w')
center_text_check.grid(row=0, column=3, padx=(2, 2), sticky='w')
autosize_col_check.grid(row=0, column=4, padx=(2, 2), sticky='w')
resize_col_check.grid(row=0, column=5, padx=(2, 2), sticky='w')
column_width_spin.grid(row=0, column=6, padx=(2, 2), sticky='w')
px_label.grid(row=0, column=7, sticky='w')

# 19 Any time a different format setting is selected, we call update_instruction to display relevant directions
format_combo.bind("<<ComboboxSelected>>", update_instruction)

# 20 Call these to ensure instructions and checkbutton are displayed correctly initially. Go to top for coment #21
update_instruction()

handle_resize_checkbutton()

# 2: This starts the event loop of the user interface. It's at the end because we need to configure our interface before we run it. It continually checks for input on the user interface. #3 is on line 181
root.mainloop()