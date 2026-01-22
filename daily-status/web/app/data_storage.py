# data_storage.py
import json
from tkinter import messagebox
import openpyxl
from reportlab.lib.pagesizes import letter
from .utils import DATA_FILE, EXCEL_FILE
from datetime import date, datetime
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import calendar
from openpyxl.styles import Border, Side, PatternFill, Alignment, Font

# Global variables
students = {}
times = []
column1_options = {}
column2_options = {}
column3_options = {}
column4_options = {}
colors = {}
header_font = Font(name="Century Gothic", size=11)  # Font for column A and rows 1 & 2
default_font = Font(name="Century Gothic", size=8)  # Font for all other cells

def initialize_data():
    print("...in initialize_data")
    """Load data from the JSON file into global variables."""
    global students, times, column1_options, column2_options, column3_options, column4_options, colors
    try:
        with open(DATA_FILE, "r") as file:
            data = json.load(file)
            students = data.get("students", {})  # Load as a dictionary
            times = data.get("times", [])
            column1_options = data.get("column1_options", {})  # Load as a dictionary
            column2_options = data.get("column2_options", {})  # Load as a dictionary
            column3_options = data.get("column3_options", {})  # Load as a dictionary
            column4_options = data.get("column4_options", {})  # Load as a dictionary
            colors = data.get("colors", {})

    except FileNotFoundError:
        print(f"{DATA_FILE} not found. Using default values.")
        students = {}
        times = []
        column1_options = {}
        column2_options = {}
        column3_options = {}
        column4_options = {}
        colors = {}

def save_data():
    data = {
        "students": students,
        "times": times,
        "column1_options": column1_options,
        "column2_options": column2_options,
        "column3_options": column3_options,
        "column4_options": column4_options,
        "colors": colors,
    }
    update_data(data)

def update_data(data):
    try:
        with open(DATA_FILE, "w") as json_file:
            json.dump(data, json_file, indent=4)
        print(f"Data successfully saved to {DATA_FILE}")
    except Exception as e:
        print(f"Error saving data to {DATA_FILE}: {e}")

def set_font(sheet, font_name="Century Gothic", header_font_size=11, default_font_size=8):
    """Set the font for all cells in the sheet with specific sizes for headers and column A."""
    header_font = Font(name=font_name, size=header_font_size)  # Font for column A and rows 1 & 2
    default_font = Font(name=font_name, size=default_font_size)  # Font for all other cells

    # Iterate through all rows and columns in the sheet
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            if cell.value is not None:  # Only apply font to non-empty cells
                # Apply header font to column A and rows 1 & 2
                if cell.row <= 3 or cell.column == 1:
                    cell.font = header_font
                else:
                    cell.font = default_font
                    
def get_weekdays_and_weekends(year, month):
    # Get the total number of days in the month
    _, num_days = calendar.monthrange(year, month)

    weekdays = []  # List to store Monday-Friday dates
    weekends = []  # List 
    
    # Map day numbers to single initials
    day_initials = ["M", "T", "W", "R", "F", " ", " "]

    # Iterate through all days of the month
    for day in range(1, num_days + 1):
        day_of_week = date(year, month, day).weekday()  # 0 = Monday, 6 = Sunday
        day_initial = day_initials[day_of_week]  # Get the single initial for the day

        if day_of_week < 5:  # Monday-Friday
            weekdays.append((day, day_initial))
        else:  # Saturday-Sunday
            weekends.append((day, day_initial))

    return weekdays, weekends

from openpyxl.styles import Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from datetime import date
import calendar

def shade_weekends(sheet, year, month):
    """Shade weekend columns, adjust column sizes, and write day initials and numbers."""
    # Define fill colors for shading
    weekend_fill = PatternFill(start_color="dbdbdb", end_color="dbdbdb", fill_type="solid")  # Dark gray for weekends
    weekday_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # White for weekdays

    # Define border style
    thin_border = Border(
        left=Side(style="thin", color="A9A9A9"),
        right=Side(style="thin", color="A9A9A9"),
        top=Side(style="thin", color="A9A9A9"),
        bottom=Side(style="thin", color="A9A9A9"),
    )

    # Map day numbers to single initials
    day_initials = ["M", "T", "W", "R", "F", " ", " "]  # Monday-Sunday

    # Get weekdays and weekends for the given month and year
    _, num_days = calendar.monthrange(year, month)

    # Write headers for each day of the month
    for day in range(1, num_days + 1):
        col_idx = day + 1  # Assuming column 1 is reserved for row labels
        col_letter = get_column_letter(col_idx)  # Get the column letter (e.g., B, C, etc.)
        day_of_week = date(year, month, day).weekday()  # 0 = Monday, 6 = Sunday

        # Write the letter day of the week in row 2
        sheet.cell(row=2, column=col_idx, value=day_initials[day_of_week])
        sheet.cell(row=2, column=col_idx).alignment = Alignment(horizontal="center", vertical="center")
        
        # Write the number day of the month in row 3
        sheet.cell(row=3, column=col_idx, value=day)
        sheet.cell(row=3, column=col_idx).alignment = Alignment(horizontal="center", vertical="center")
        
        # Set column width based on whether the day is a weekend or weekday
        if day_of_week >= 5:  # Saturday or Sunday
            sheet.column_dimensions[col_letter].width = 3  # Weekend column width (24 pixels)
        else:
            sheet.column_dimensions[col_letter].width = 9  # Weekday column width (75 pixels)

        # Apply shading and borders to all rows in the column
        for row_idx in range(2, sheet.max_row + 1):  # Start from row 2
            cell = sheet.cell(row=row_idx, column=col_idx)
            if day_of_week >= 5:  # Saturday or Sunday
                cell.fill = weekend_fill  # Apply weekend fill color
            else:
                cell.fill = weekday_fill  # Apply weekday fill color
                # Iterate through all rows and columns in the sheet
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.border = thin_border  # Apply borders to all cells    

def write_to_excel(data):
    global times
    print("...in write_to_excel")
    print(f"Data being written: {data}")
    # print(f"Looking for time {data['Time']} in sheet {sheet_name}")

    """Write data to an Excel file with student-specific sheets for each month."""
    # Get the current date
    today = datetime.now()
    current_month = today.strftime("%B")  # Full month name (e.g., "December")
    current_day = today.day  # Day of the month (e.g., 31)

    # Load the workbook if it exists, otherwise create a new one
    if os.path.exists(EXCEL_FILE):
        workbook = openpyxl.load_workbook(EXCEL_FILE)
    else:
        workbook = openpyxl.Workbook()
        # Remove the default sheet created by openpyxl
        default_sheet = workbook.active
        workbook.remove(default_sheet)

    print("found workbook")

    # Get the student's name
    student_name = data["Student"]

    # Create the sheet name for the current student and month
    sheet_name = f"{student_name}-{current_month}"

    print(f"Looking for time {data['Time']} in sheet {sheet_name}")

    # Check if the sheet already exists, otherwise create it
    if sheet_name not in workbook.sheetnames:
        print ("sheet not in workbook - create")
        sheet = workbook.create_sheet(title=sheet_name)

        # Add the student's name in the first row
        sheet["A1"] = student_name
        sheet["A1"].alignment = Alignment(horizontal="left", vertical="center")
        sheet["B1"] = current_month
        sheet["B1"].alignment = Alignment(horizontal="left", vertical="center")
        sheet["C1"] = today.year
        sheet["C1"].alignment = Alignment(horizontal="left", vertical="center")

        # Add headers in the third row
        sheet["A3"] = "Dates"
        # for day in range(1, 32):  # Maximum days in a month
        for day in range(1, calendar.monthrange(today.year, today.month)[1] + 1):
            col_letter = openpyxl.utils.get_column_letter(day + 1)  # Start from column B
            sheet[f"{col_letter}3"] = str(day)

        print("Times list:", times)
        # Add times in column A starting from row 3
        for idx, time in enumerate(times, start=4):
            sheet[f"A{idx}"] = time

        # Apply alternate shading
        shade_weekends(sheet, year=today.year, month=today.month)
        set_font(sheet)

    else:
        # If the sheet already exists, load it
        print("sheet already exists")
        sheet = workbook[sheet_name]

    for idx, time in enumerate(times, start=4):
        sheet[f"A{idx}"] = time

    # Find the column for the current day
    day_column = None
    for col in range(2, sheet.max_column + 1):  # Start from column B
        col_letter = openpyxl.utils.get_column_letter(col)
        cell_value = sheet[f"{col_letter}3"].value  # Get the value in row 3 for the current column

        # Ensure both values are integers for comparison
        if cell_value is not None and int(cell_value) == int(current_day):
            print("found match day column")
            day_column = col_letter
            break

    # Raise an error if the column is not found
    if day_column is None:
        raise ValueError(f"Could not find the column for day {current_day} in sheet {sheet.title}.")

    # Find the row for the selected time
    time_row = None
    for row in range(3, sheet.max_row + 1):  # Start from row 3
        if sheet[f"A{row}"].value == data["Time"]:
            time_row = row
            break

    if not time_row:
        raise ValueError(f"Could not find the row for time {data['Time']} in sheet {sheet_name}.")

    print("time_row:", time_row)
    print(data)

    # Collect values from column1, column2, column3, column4, and notes
    values_to_concatenate = [
        data[key].strip() for key in ["column2", "column1", "column3", "column4"]
        if key in data and data[key] and data[key].strip() != "UNSELECTED"  # Exclude empty, None, or "UNSELECTED"
    ]
    print("values_to_concatenate:", values_to_concatenate)

    # Join the remaining values into a single string with a newline delimiter
    concatenated_values = "".join(values_to_concatenate)

    # Add "Notes" on a new line if it exists
    if "Notes" in data and data["Notes"] and data["Notes"].strip() != "UNSELECTED":
        concatenated_values += "\n" + data["Notes"].strip()  # Add Notes on a new line

    username = data.get("Username", "N/A")
    # Concatenate the username with the other values
    concatenated_values = "(" + username + ") \n" + concatenated_values

    # Write the concatenated values to the appropriate cell
    sheet[f"{day_column}{time_row}"] = concatenated_values
    sheet[f"{day_column}{time_row}"].font = default_font
    enable_text_wrapping(sheet, wrap_notes_only=True)
    # adjust_column_width(sheet)

    # Save the workbook
    workbook.save(EXCEL_FILE)
    # enable_text_wrapping(EXCEL_FILE)

    print("done with excel")

    return True

def enable_text_wrapping(sheet, wrap_notes_only=False):
    # Iterate through all cells in the sheet
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value:
                if wrap_notes_only:
                    # Enable wrapping only if the cell contains a newline (Notes)
                    if "\n" in str(cell.value):
                        cell.alignment = Alignment(wrap_text=True)
                else:
                    # Enable wrapping for all cells
                    cell.alignment = Alignment(wrap_text=True)

def adjust_column_width(sheet):
    """Adjust the column width to fit the text in each column, skipping empty columns."""
    for col in range(1, sheet.max_column + 1):  # Iterate through all columns
        col_letter = get_column_letter(col)  # Get the column letter (e.g., A, B, C)
        max_length = 0  # Initialize the maximum length for the column
        has_data = False  # Flag to check if the column has any data

        for row in range(1, sheet.max_row + 1):  # Iterate through all rows in the column
            cell = sheet.cell(row=row, column=col)
            if cell.value:  # Check if the cell has a value
                has_data = True  # Mark that the column has data
                # Split the cell value into lines based on wrapping (newlines)
                lines = str(cell.value).split("\n")
                # Find the longest line in the cell
                max_line_length = max(len(line) for line in lines)
                # Update the maximum length for the column
                max_length = max(max_length, max_line_length)

        # Only adjust the column width if the column has data
        if has_data:
            # Set the column width (add some padding for better readability)
            sheet.column_dimensions[col_letter].width = max_length + 2