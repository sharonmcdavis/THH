# data_storage.py
import json
from tkinter import messagebox
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from .utils import DATA_FILE, EXCEL_FILE, PDF_FILE

# Global variables
students = {}
times = []
column1_options = {}
column2_options = {}
column3_options = {}
column4_options = {}
colors = {}

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
        students = dict(sorted(students.items(), key=lambda x: x[0], reverse=False))

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

# Function to write data to an Excel sheet
import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime
import os
from openpyxl import load_workbook
from fpdf import FPDF
from openpyxl.styles import PatternFill

def apply_alternate_shading(sheet):
    print("...in apply_alternate_shading")
    """Apply alternate shading to every 5 columns in the Excel sheet."""
    # Define fill colors for shading
    fill_color_1 = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")  # Light gray
    fill_color_2 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # White (no shading)

    # Iterate through columns in groups of 5
    for col_idx in range(2, sheet.max_column + 1):
        # Determine the shading group (alternate every 5 columns)
        shading_group = (col_idx - 1) // 4
        fill_color = fill_color_1 if shading_group % 2 == 0 else fill_color_2

        # Apply the fill color to all rows in the column
        for row_idx in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.fill = fill_color

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
        sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

        # Add headers in the second row
        sheet["A2"] = "Dates"
        for day in range(1, 32):  # Maximum days in a month
            col_letter = openpyxl.utils.get_column_letter(day + 1)  # Start from column B
            sheet[f"{col_letter}2"] = str(day)

        print("Times list:", times)
        # Add times in column A starting from row 3
        for idx, time in enumerate(times, start=3):
            sheet[f"A{idx}"] = time

        # Apply alternate shading
        apply_alternate_shading(sheet)

    else:
        # If the sheet already exists, load it
        print("sheet already exists")
        sheet = workbook[sheet_name]

    print("sheet: ", sheet)
    print("data['Time']:", data["Time"])
    for idx, time in enumerate(times, start=3):
        sheet[f"A{idx}"] = time
        print(f"Writing time '{time}' to A{idx}")

    # Find the column for the current day
    day_column = None
    for col in range(2, sheet.max_column + 1):  # Start from column B
        col_letter = openpyxl.utils.get_column_letter(col)
        if sheet[f"{col_letter}2"].value == str(current_day):
            day_column = col_letter
            break

    if not day_column:
        raise ValueError(f"Could not find the column for day {current_day} in sheet {sheet_name}.")
    print("day_col:", day_column)

    # Find the row for the selected time
    time_row = None
    for row in range(3, sheet.max_row + 1):  # Start from row 3
        print("row: ", row)
        print("data[time]:", data["Time"])
        if sheet[f"A{row}"].value == data["Time"]:
            time_row = row
            break

    if not time_row:
        raise ValueError(f"Could not find the row for time {data['Time']} in sheet {sheet_name}.")

    print("time_row:", time_row)
    print(data)

    # Collect values from column1, column2, column3, column4, and notes
    values_to_concatenate = [
        data[key].strip() for key in ["column1", "column2", "column3", "column4", "Notes"]
        if key in data and data[key] and data[key].strip() != "UNSELECTED"  # Exclude empty, None, or "UNSELECTED"
    ]

    # Join the remaining values into a single string with a newline delimiter
    concatenated_values = "\n".join(values_to_concatenate)

    print("concat value:", concatenated_values)
    username = data.get("Username", "N/A")
    print("username:", username)
    # Concatenate the username with the other values
    concatenated_values = "(" + username + ") " + concatenated_values
    print("concat value:", concatenated_values)

    # Write the concatenated values to the appropriate cell
    sheet[f"{day_column}{time_row}"] = concatenated_values

    # Save the workbook
    workbook.save(EXCEL_FILE)
    enable_text_wrapping(EXCEL_FILE)

    print("done with excel")

    return True

from openpyxl import load_workbook
from openpyxl.styles import Alignment

def enable_text_wrapping(excel_file):
    print("...in enable_text_wrapping")
    # Load the workbook
    workbook = load_workbook(excel_file)

    # Iterate through all sheets in the workbook
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Iterate through all cells in the sheet
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and "\n" in str(cell.value):  # Check if the cell contains a newline
                    cell.alignment = Alignment(wrap_text=True)  # Enable text wrapping

    # Save the workbook
    workbook.save(excel_file)
    print(f"Text wrapping enabled in {excel_file}")
