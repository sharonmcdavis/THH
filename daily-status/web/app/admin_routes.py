from tkinter import messagebox
from flask import Blueprint, render_template, request, redirect, session, url_for, flash, send_file
from fpdf import FPDF
import openpyxl
from .data_storage import save_data, update_data, EXCEL_FILE, PDF_FILE
from .data_loader import load_data_from_file
from .utils import login_required, ADMIN_PASSWORD, admin_login_required
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

admin = Blueprint('admin', __name__)

# Load data from the JSON file
data = load_data_from_file()

@admin.route('/')
@login_required
@admin_login_required
def admin_login():
    """Render the main page."""
    return render_template(
        'admin_login.html'
    )

@admin.route('/verify_admin', methods=['POST'])
@login_required
@admin_login_required
def verify_admin():
    # Get the password entered by the user
    entered_password = request.form.get('admin_password')

    # Verify the password
    if entered_password == ADMIN_PASSWORD:
        session['admin_logged_in'] = True
        return redirect(url_for('admin.admin_window'))  # Redirect to the admin page
    else:
        flash("Incorrect password. Access denied.", "error")
        return redirect(url_for('admin.admin_login'))  # Redirect back to the main page
    
@admin.route('/open')
@login_required
@admin_login_required
def admin_window():
    """Render the main page."""
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/add_student', methods=['POST'])
@login_required
@admin_login_required
def add_student():
    from .data_storage import students
    """Add a new student."""
    student = request.form.get('student')
    print("student:", student)

    if student and student.lower not in [s.lower() for s in students]:
        students.append(student)
        save_data()
        flash(f"Student '{student}' added successfully!", "success")
    else:
        flash("Student already exists or is invalid.", "error")
    data = load_data_from_file()
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/remove_student', methods=['POST'])
@login_required
@admin_login_required
def remove_student():
    from .data_storage import students
    """Remove a student."""
    student = request.form.get('student')
    if student and student in students:
        students.remove(student)
        save_data()
        flash(f"Student '{student}' removed successfully!", "success")
    else:
        flash("Student not found.", "error")
    data = load_data_from_file()
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/add_time', methods=['POST'])
@login_required
@admin_login_required
def add_time():
    from .data_storage import times
    """Add a new time."""
    time = request.form.get('time')
    if time and time not in times:
        times.append(time)
        save_data()
        flash(f"Time '{time}' added successfully!", "success")
    else:
        flash("Time already exists or is invalid.", "error")
    data = load_data_from_file()
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/remove_time', methods=['POST'])
@login_required
@admin_login_required
def remove_time():
    from .data_storage import times
    """Remove a time."""
    time = request.form.get('time')
    if time and time in times:
        times.remove(time)
        save_data()
        flash(f"Time '{time}' removed successfully!", "success")
    else:
        flash("Time not found.", "error")
    data = load_data_from_file()
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/add_column', methods=['POST'])
@login_required
@admin_login_required
def add_column():
    form_data = request.form.to_dict()
    print("form_data:", form_data)

    # Extract column name, key, and value
    column_name = form_data.get('column_name')
    key = form_data.get('key')
    value = form_data.get('value')

    # Load the current data
    data = load_data_from_file()

   # Ensure the column exists in the data
    if column_name not in data:
        flash(f"Column '{column_name}' does not exist.", "error")
        return redirect(url_for('admin.admin_window'))

    # Add the key-value pair to the specified column
    if key in data[column_name]:
        flash(f"Key '{key}' already exists in {column_name}.", "error")
    else:
        data[column_name][key] = value
        update_data(data)  # Save the updated data
        flash(f"Key '{key}' with value '{value}' added.", "success")

    data = load_data_from_file()
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/remove_column', methods=['POST'])
@login_required
@admin_login_required
def remove_column():
    form_data = request.form.to_dict()

    # Process the form data (should only contain one key-value pair)
    if len(form_data) != 1:
        flash("Invalid form submission.", "error")
        return redirect(url_for('admin.admin_window'))

    # Extract the column name and value
    column_name, value_to_remove = next(iter(form_data.items()))
    print(f"Column: {column_name}, Value: {value_to_remove}")

    data = load_data_from_file()
    print("data:", data)
    
    # Ensure the column exists in the data
    if column_name in data and value_to_remove in data[column_name]:
        # Remove the value from the column
        del data[column_name][value_to_remove]
        print("data after:", data)
        update_data(data)  # Save the updated data
        flash(f"Value '{value_to_remove}' removed.", "success")
    else:
        flash(f"Value '{value_to_remove}' not found.", "error")

    data = load_data_from_file()
    return render_template(
        'admin.html',
        data=data
    )

@admin.route('/excel', methods=['POST'])
@login_required
@admin_login_required
def excel():
    # Check if the file exists
    if os.path.exists(EXCEL_FILE):
        try:
            print("Opening Excel file...")
            os.startfile(EXCEL_FILE)  # Opens the file with the default application
            # flash(f"Excel file '{EXCEL_FILE}' opened successfully.", "success")
        except Exception as e:
            print(f"Error opening Excel file: {e}")
            flash(f"An error occurred while opening the Excel file: {e}", "error")
    else:
        flash("Excel file not found.", "error")
    
    return send_file(EXCEL_FILE, as_attachment=True)

class PDF(FPDF):
    def __init__(self, orientation="L", unit="mm", format="A4"):
        super().__init__(orientation, unit, format)
        self.set_auto_page_break(auto=True, margin=10)

    def add_table(self, data, page_width, col_headers=None):
        self.set_font("Arial", size=8)

        # Calculate column widths dynamically
        num_columns = len(data[0]) if data else 1
        col_width = page_width / num_columns

        # Add column headers if provided
        if col_headers:
            for header in col_headers:
                self.cell(col_width, 10, header, border=1, align="C")
            self.ln()

        # Add rows of data
        for row in data:
            for cell in row:
                self.cell(col_width, 10, str(cell) if cell is not None else "", border=1, align="C")
            self.ln()

@admin.route('/pdf', methods=['POST'])
@login_required
@admin_login_required
def pdf():
    from .data_storage import EXCEL_FILE, PDF_FILE
    from openpyxl import load_workbook
    from fpdf import FPDF

    if not os.path.exists(EXCEL_FILE):
        flash("Excel file not found.", "error")
        return redirect(url_for('admin.admin_window'))

    # Load the workbook
    workbook = load_workbook(EXCEL_FILE)

    # Create a PDF object
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=10)

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Split columns into three groups
        columns_group_1 = list(sheet.iter_cols(min_col=1, max_col=11, values_only=True))  # A-K
        columns_group_2 = [list(sheet.iter_cols(min_col=1, max_col=1, values_only=True))[0]]  # A (repeated)
        columns_group_2 += list(sheet.iter_cols(min_col=12, max_col=21, values_only=True))  # L-U
        columns_group_3 = [list(sheet.iter_cols(min_col=1, max_col=1, values_only=True))[0]]  # A (repeated)
        columns_group_3 += list(sheet.iter_cols(min_col=22, max_col=32, values_only=True))  # V-AF

        # Add the first page for columns A-K
        pdf.add_page()
        pdf.cell(0, 10, f"{sheet_name} (1-10)", ln=True, align="C")
        pdf.ln(5)
        add_columns_to_pdf(pdf, columns_group_1)

        # Add the second page for column A and columns L-U
        pdf.add_page()
        pdf.cell(0, 10, f"{sheet_name} (11-20)", ln=True, align="C")
        pdf.ln(5)
        add_columns_to_pdf(pdf, columns_group_2)

    # Add the third page for column A and columns V-AF
    pdf.add_page()
    pdf.cell(0, 10, f"{sheet_name} (21-30/31)", ln=True, align="C")
    pdf.ln(5)
    add_columns_to_pdf(pdf, columns_group_3)

    # Save the PDF file
    pdf.output(PDF_FILE)
    print(f"PDF saved to {PDF_FILE}")
    # os.startfile(PDF_FILE)

    if not os.path.exists(PDF_FILE):
        flash("PDF file not found.", "error")
        return
    return send_file(PDF_FILE, as_attachment=True)


def add_columns_to_pdf(pdf, columns):
    """
    Add columns of data to the PDF.
    Each column is treated as a row in the PDF.
    Dynamically wrap text to fit within the cell and ensure proper alignment.
    """
    # Transpose the columns to rows for easier processing
    rows = list(zip(*columns))

    # Calculate column width dynamically
    page_width = pdf.w - 20  # Account for margins
    col_width = page_width / len(columns)

    # Add rows to the PDF
    for row in rows:
        # Calculate the height of the tallest cell in the row
        row_heights = []
        wrapped_cells = []
        for cell in row:
            text = str(cell) if cell is not None else ""
            wrapped_text, font_size = wrap_text_to_fit(text, col_width, pdf)
            wrapped_cells.append((wrapped_text, font_size))
            row_heights.append(10 * len(wrapped_text.split("\n")))  # 10 units per line

        max_row_height = max(row_heights)

        # Check if the row fits on the current page
        if pdf.get_y() + max_row_height > pdf.h - 15:  # Account for bottom margin
            pdf.add_page()  # Add a new page if the row doesn't fit

        # Write each cell in the row
        y_start = pdf.get_y()  # Get the starting y-coordinate of the row
        for i, (cell_text, font_size) in enumerate(wrapped_cells):
            x_start = pdf.get_x()  # Get the starting x-coordinate of the cell
            pdf.set_font("Arial", size=font_size)  # Set the font size for the cell
            pdf.multi_cell(col_width, 10, cell_text, border=1, align="C")
            pdf.set_xy(x_start + col_width, y_start)  # Move to the next cell in the row

        # Move to the next row
        pdf.set_y(y_start + max_row_height)

def wrap_text_to_fit(text, col_width, pdf):
    """
    Wrap text to fit within the cell width by breaking it into multiple lines.
    Dynamically adjust font size for multi-line text.
    """
    max_font_size = 10  # Default font size
    min_font_size = 6   # Minimum font size for multi-line text
    wrapped_lines = []
    current_line = ""

    # Split the text into words
    words = text.split(" ")

    for word in words:
        # Check if adding the word exceeds the column width
        if len(current_line) + len(word) + 1 <= 10 and pdf.get_string_width(current_line + " " + word) <= col_width - 2:
            current_line += " " + word if current_line else word
        else:
            wrapped_lines.append(current_line)  # Add the current line to the wrapped lines
            current_line = word  # Start a new line with the current word

    # Add the last line
    if current_line:
        wrapped_lines.append(current_line)

    # Adjust font size based on the number of lines
    num_lines = len(wrapped_lines)
    font_size = max(min_font_size, max_font_size - (num_lines - 1))

    # Join the wrapped lines with newline characters
    return "\n".join(wrapped_lines), font_size