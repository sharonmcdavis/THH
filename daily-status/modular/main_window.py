import tkinter as tk
from utils import toggle_button
from admin_window import open_admin_window

def create_main_window(root, students, times, column1, column2, column3, column4):
    # Variables to track selected buttons
    selected_student = tk.StringVar()
    selected_time = tk.StringVar()
    column1_values = {label: tk.BooleanVar() for label in column1}
    column2_values = {label: tk.BooleanVar() for label in column2}
    column3_values = {label: tk.BooleanVar() for label in column3}
    column4_values = {label: tk.BooleanVar() for label in column4}

    # Function to refresh the main window
    def refresh_main_window():
        # Clear and repopulate the main window's widgets with updated data
        for widget in root.winfo_children():
            widget.destroy()
        create_main_window(root, students, times, column1, column2, column3, column4)

    # Add a button to open the admin window
    tk.Button(root, text="Admin Functions", command=lambda: open_admin_window(root, students, times, column1, column2, column3, column4, refresh_main_window)).pack(pady=10)

    # Example: Add toggle buttons for students
    for student in students:
        tk.Button(root, text=student, command=lambda s=student: toggle_button(selected_student, s)).pack(pady=5)

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
        "Notes": notes_text.get("1.0", "end-1c").strip(),  
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
        sheet.append(["Student", "Time", "Column1", "Column2", "Column3", "Column4", "Notes"])

    # Write the data
    sheet.append([
        data["Student"],
        data["Time"],
        ", ".join(data["Column1"]),
        ", ".join(data["Column2"]),
        ", ".join(data["Column3"]),
        ", ".join(data["Column4"]),
        data["Notes"],  # Add the notes to the Excel file
    ])
    workbook.save(file_name)

     