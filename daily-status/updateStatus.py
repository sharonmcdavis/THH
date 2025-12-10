import tkinter as tk
from tkinter import messagebox
import openpyxl
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from tkinter import ttk  # Import ttk for the Separator widget

# Create the main application window
root = tk.Tk()
root.title("Student Activity Tracker")
root.geometry("1200x800")  # Set the window size

# Define the students and times
students = ["Silas", "Journee", "Nate", "Rip", "Zi", "Cami", "Braxton", "Arturo", "Emmy"]
times = ["7:45", "8:30", "9:00", "9:30", "10:00", "10:30", "11:00", "11:30", "12:00", "12:30", "1:00", "1:30", "2:00", "AC"]

# Define the button labels for the columns
column1 = ["W - wet", "P - dirty", "D - dry"]
column2 = ["+ stood and peed", "- stood and did nothing"]
column3 = ["S+ sat and peed", "S++ sat and pooped", "S- sat and did nothing"]
column4 = ["C+ peed in cup", "C- stood but did not pee in cup", "ACC - had an accident in underwear"]

# Variables to track selected buttons
selected_student = tk.StringVar()
selected_time = tk.StringVar()
column1_values = {label: tk.BooleanVar() for label in column1}
column2_values = {label: tk.BooleanVar() for label in column2}
column3_values = {label: tk.BooleanVar() for label in column3}
column4_values = {label: tk.BooleanVar() for label in column4}

# Function to handle toggle button clicks
def toggle_button(var, value):
    if var.get() == value:
        var.set("")  # Deselect if clicked again
    else:
        var.set(value)

# Function to submit the data
def submit_data():
    if not selected_student.get():
        messagebox.showerror("Error", "Please select a student.")
        return
    if not selected_time.get():
        messagebox.showerror("Error", "Please select a time.")
        return

    # Collect selected values
    data = {
        "Student": selected_student.get(),
        "Time": selected_time.get(),
        "Column1": [label for label, var in column1_values.items() if var.get()],
        "Column2": [label for label, var in column2_values.items() if var.get()],
        "Column3": [label for label, var in column3_values.items() if var.get()],
        "Column4": [label for label, var in column4_values.items() if var.get()],
    }

    # Write data to Excel
    write_to_excel(data)
    messagebox.showinfo("Success", "Data submitted successfully!")
    reset_buttons()

# Function to write data to an Excel sheet
def write_to_excel(data):
    file_name = "student_activity.xlsx"
    try:
        workbook = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Activity Log"

    # Write headers if the sheet is empty
    if sheet.max_row == 1:
        sheet.append(["Student", "Time", "Column1", "Column2", "Column3", "Column4"])

    # Write the data
    sheet.append([
        data["Student"],
        data["Time"],
        ", ".join(data["Column1"]),
        ", ".join(data["Column2"]),
        ", ".join(data["Column3"]),
        ", ".join(data["Column4"]),
    ])
    workbook.save(file_name)

# Function to reset all buttons
def reset_buttons():
    selected_student.set("")
    selected_time.set("")
    for var in column1_values.values():
        var.set(False)
    for var in column2_values.values():
        var.set(False)
    for var in column3_values.values():
        var.set(False)
    for var in column4_values.values():
        var.set(False)

# Function to open the Excel file
def open_excel_file():
    file_name = "student_activity.xlsx"
    if os.path.exists(file_name):
        os.startfile(file_name)  # Opens the file with the default application
    else:
        messagebox.showerror("Error", "Excel file not found.")

# Function to convert Excel to PDF and open it
def convert_to_pdf():
    excel_file = "student_activity.xlsx"
    pdf_file = "student_activity.pdf"

    if not os.path.exists(excel_file):
        messagebox.showerror("Error", "Excel file not found.")
        return

    # Create a PDF
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    pdf = canvas.Canvas(pdf_file, pagesize=letter)
    pdf.setFont("Helvetica", 10)

    # Write data from Excel to PDF
    x, y = 50, 750  # Starting position
    for row in sheet.iter_rows(values_only=True):
        line = " | ".join([str(cell) if cell is not None else "" for cell in row])
        pdf.drawString(x, y, line)
        y -= 20
        if y < 50:  # Create a new page if the content exceeds the page
            pdf.showPage()
            pdf.setFont("Helvetica", 10)
            y = 750

    pdf.save()

    # Open the PDF
    os.startfile(pdf_file)


# Vertical Students (aligned to the left)
student_frame = tk.Frame(root)
student_frame.pack(side=tk.LEFT, padx=10, pady=10, anchor="n")  # Align to the left
tk.Label(student_frame, text="Students").pack(pady=5)
for student in students:
    btn = tk.Radiobutton(
        student_frame, text=student, variable=selected_student, value=student,
        indicatoron=False, width=15, height=2, bg="lightblue", selectcolor="blue"
    )
    btn.pack(pady=2)

# Vertical Times (aligned to the left)
time_frame = tk.Frame(root)
time_frame.pack(side=tk.LEFT, padx=10, pady=10, anchor="n")  # Align to the left
tk.Label(time_frame, text="Times").pack(pady=5)
for time in times:
    btn = tk.Radiobutton(
        time_frame, text=time, variable=selected_time, value=time,
        indicatoron=False, width=15, height=2, bg="lightgreen", selectcolor="green"
    )
    btn.pack(pady=2)

# Create the table with toggle buttons
table_frame = tk.Frame(root)
table_frame.pack(pady=25, anchor="w")  # Align the entire table to the left

# Column 1
col1_frame = tk.Frame(table_frame)
col1_frame.grid(row=0, column=0, padx=10, sticky="w")  # Position next to the "time" column
for label in column1:
    btn = tk.Checkbutton(
        col1_frame, text=label, variable=column1_values[label],
        width=30, height=1, anchor="w"  # Align text and checkbox to the left
    )
    btn.pack(anchor="w", pady=2)

# Add a horizontal line (separator) after Column 1 options
separator = ttk.Separator(table_frame, orient="horizontal")
separator.grid(row=1, column=0, sticky="ew", pady=10)  # Stretch the line horizontally

# Column 2
col2_frame = tk.Frame(table_frame)
col2_frame.grid(row=2, column=0, padx=10, sticky="w")  # Position next to Column 1
for label in column2:
    btn = tk.Checkbutton(
        col2_frame, text=label, variable=column2_values[label],
        width=30, height=1, anchor="w"  # Align text and checkbox to the left
    )
    btn.pack(anchor="w", pady=2)

# Add a horizontal line (separator) after Column 1 options
separator = ttk.Separator(table_frame, orient="horizontal")
separator.grid(row=3, column=0, sticky="ew", pady=10)  # Stretch the line horizontally

# Column 3
col3_frame = tk.Frame(table_frame)
col3_frame.grid(row=4, column=0, padx=10, sticky="w")  # Position next to Column 2
for label in column3:
    btn = tk.Checkbutton(
        col3_frame, text=label, variable=column3_values[label],
        width=30, height=1, anchor="w"  # Align text and checkbox to the left
    )
    btn.pack(anchor="w", pady=2)

# Add a horizontal line (separator) after Column 1 options
separator = ttk.Separator(table_frame, orient="horizontal")
separator.grid(row=5, column=0, sticky="ew", pady=10)  # Stretch the line horizontally

# Column 4
col4_frame = tk.Frame(table_frame)
col4_frame.grid(row=6, column=0, padx=10, sticky="w")  # Position next to Column 3
for label in column4:
    btn = tk.Checkbutton(
        col4_frame, text=label, variable=column4_values[label],
        width=30, height=1, anchor="w"  # Align text and checkbox to the left
    )
    btn.pack(anchor="w", pady=2)

button_frame = tk.Frame(root)
button_frame.pack(padx=50, pady=25, anchor="w")  # Align the entire table to the left

submit_button = tk.Button(button_frame, text="Submit", command=submit_data, bg="orange", height=2, width=20)
submit_button.pack(pady=5)  # Add vertical spacing between buttons

view_excel_button = tk.Button(button_frame, text="View Excel", command=open_excel_file, bg="blue", fg="white", height=2, width=20)
view_excel_button.pack(pady=5)  # Add vertical spacing between buttons

view_pdf_button = tk.Button(button_frame, text="View PDF", command=convert_to_pdf, bg="green", fg="white", height=2, width=20)
view_pdf_button.pack(pady=5)  # Add vertical spacing between buttons

# Run the application
root.mainloop()