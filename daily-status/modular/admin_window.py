import tkinter as tk
from tkinter import messagebox
from data_manager import save_data
import openpyxl

def open_admin_window(root, students, times, column1, column2, column3, column4, refresh_main_window):
    admin_window = tk.Toplevel(root)
    admin_window.title("Admin Functions")
    admin_window.geometry("600x500")

    # Make the admin window modal
    admin_window.transient(root)
    admin_window.grab_set()
    admin_window.focus_set()

    # Create a scrollable frame
    canvas = tk.Canvas(admin_window)
    scrollbar = tk.Scrollbar(admin_window, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Add Exit Button to Admin Window
    tk.Button(scrollable_frame, text="Exit Admin", command=admin_window.destroy, bg="red", fg="white").grid(row=20, column=0, pady=10, sticky="w")
    
    # Add the "Manage Column Options" section
    tk.Label(scrollable_frame, text="Manage Column Options", font=("Arial", 12, "bold")).grid(row=6, column=0, sticky="w", pady=5)

    # Dropdown for selecting the column
    column_var = tk.StringVar(value="Column1")
    column_dropdown = tk.OptionMenu(scrollable_frame, column_var, "Students, Times, Column1", "Column2", "Column3", "Column4", command=lambda _: update_listbox(column_listbox, get_column_list(column_var.get())))
    column_dropdown.grid(row=7, column=0, sticky="w", pady=5)

    # Listbox for displaying current column options
    column_listbox = tk.Listbox(scrollable_frame, height=6, width=30)
    column_listbox.grid(row=8, column=0, sticky="w", pady=5)
    update_listbox(column_listbox, column1)  # Default to Column1

    # Entry field for adding/modifying options
    option_entry = tk.Entry(scrollable_frame, width=30)
    option_entry.grid(row=9, column=0, sticky="w", pady=5)

    # "Add Option" button directly to the right of the entry field
    tk.Button(scrollable_frame, text="Add Option", command=lambda: add_item(option_entry, get_column_list(column_var.get()), column_listbox), bg="green", fg="white").grid(row=9, column=1, sticky="w", padx=5)

    # Frame for "Modify" and "Remove" buttons, aligned with the top of the listbox
    column_button_frame = tk.Frame(scrollable_frame)
    column_button_frame.grid(row=8, column=1, sticky="n", padx=5)  # Align with the top of the listbox
    tk.Button(column_button_frame, text="Modify Option", command=lambda: modify_item(option_entry, get_column_list(column_var.get()), column_listbox), bg="purple", fg="white").pack(fill="x", pady=2)
    tk.Button(column_button_frame, text="Remove Option", command=lambda: remove_item(option_entry, get_column_list(column_var.get()), column_listbox), bg="light gray", fg="black").pack(fill="x", pady=2)


def save_and_close(admin_window, students, times, column1, column2, column3, column4, refresh_main_window):
    save_data(students, times, column1, column2, column3, column4)
    refresh_main_window()  # Refresh the main window with updated data
    admin_window.destroy()

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