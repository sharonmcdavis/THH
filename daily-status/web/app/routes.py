import os
import shutil
from flask import Blueprint, jsonify, render_template, request, redirect, url_for, flash, session
from datetime import datetime
from .data_storage import initialize_data, write_to_excel
from .utils import login_required, WEB_PASSWORD, EXCEL_FILE, BACKUP_FOLDER
import pandas as pd

main = Blueprint('main', __name__)

@main.route('/', methods=['GET', 'POST'])
def index():
    print("in the main default /")
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username and password == WEB_PASSWORD:
            # Save login state and timestamp in the session
            session['username'] = username
            session['logged_in'] = True
            session['login_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            create_backup()
            return redirect(url_for('main.main_window'))
        else:
            flash("Access restricted. Invalid login credentials.", "error")
            return render_template('index.html')
    return render_template('index.html')

def create_backup():
    print("create backup file")
    print("EXCEL_FILE:", EXCEL_FILE)
    try:
        # Ensure the backup folder exists
        if not os.path.exists(BACKUP_FOLDER):
            print("backup folder does NOT exist")
            os.makedirs(BACKUP_FOLDER)
            print("backup folder created")

        # Generate the daily backup filename
        today_date = datetime.now().strftime('%Y-%m-%d')
        daily_backup_file = os.path.join(BACKUP_FOLDER, f"excel_backup_{today_date}.xlsx")
        print("backup file:", daily_backup_file)

        # Check if the backup already exists
        if not os.path.exists(daily_backup_file):
            print("backup file does NOT exist - create")
            # Create the backup if it doesn't exist
            shutil.copy(EXCEL_FILE, daily_backup_file)
            print(f"Daily backup created: {daily_backup_file}")
        else:
            print(f"Backup for today already exists: {daily_backup_file}")
    except Exception as e:
        print(f"Error creating backup: {e}")

@main.route('/main')
@login_required
def main_window():
    from .data_storage import students, times, column1_options, column2_options, column3_options, column4_options
    # Get the current date and format it
    current_date = datetime.now().strftime("%B %d, %Y")  # Example: "January 8, 2026"

   # Retrieve the previously selected values from the session
    student = session.get('student', '')
    time = session.get('time', '')
    column1_selection = session.get('column1', '')
    column2_selection = session.get('column2', '')
    column3_selection = session.get('column3', '')
    column4_selection = session.get('column4', '')
    notes = session.get('notes', '')

    students = dict(sorted(students.items(), key=lambda item: item[0].casefold()))
    
    """Render the main page."""
    return render_template(
        'main.html',
        students=students,
        times=times,
        column1_options=column1_options,
        column2_options=column2_options,
        column3_options=column3_options,
        column4_options=column4_options,
        student=student,
        time=time,
        column1=column1_selection,
        column2=column2_selection,
        column3=column3_selection,
        column4=column4_selection,
        notes=notes,
        date=current_date
    )

def clear_session():
    # Clear specific session attributes
    session.pop('student', None)
    session.pop('time', None)
    session.pop('column1', None)
    session.pop('column2', None)
    session.pop('column3', None)
    session.pop('column4', None)
    session.pop('notes', None)


@main.route('/submit', methods=['POST'])
@login_required
def submit():
    print("in the main submit")
    """Handle form submission."""
    student = request.form.get('student')
    time = request.form.get('time')
    column1 = request.form.get('column1')
    column2 = request.form.get('column2')
    column3 = request.form.get('column3')
    column4 = request.form.get('column4')
    notes = request.form.get('notes')

    session['student'] = student
    session['time'] = time
    session['column1'] = column1
    session['column2'] = column2
    session['column3'] = column3
    session['column4'] = column4
    session['notes'] = notes

    print("student:", student)
    print("notes:", notes)
    
    if not student:
        flash("Please select a student.", "error")
        return redirect(url_for('main.main_window'))
    
    if not time:
        flash("Please select a time.", "error")
        return redirect(url_for('main.main_window'))
    
    # Ensure at least one column has a selection
    columns = {
        "column1": column1,
        "column2": column2,
        "column3": column3,
        "column4": column4,
    }
    selected_columns = {key: value for key, value in columns.items() if value}  # Exclude None or empty values

    if not selected_columns:
        print("column not selected - ")
        flash("Please select at least one status for the selected student.", "error")
        return redirect(url_for('main.main_window'))


    # Collect selected values for each column
    data = {
        "Username": session['username'],
        "Student": student,
        "Time": time,
        **selected_columns,
        "Notes": notes,
    }

    # Write data to Excel and refresh the main window if successful
    if write_to_excel(data):
        clear_session()
        flash("Data submitted successfully!", "success")
        return redirect(url_for('main.main_window'))
    else:
        flash("Error saving data.", "error")
        return redirect(url_for('main.main_window'))
    
@main.route('/logout', methods=['GET'])
def logout():
    print("in the logout")
    # Clear the session to remove the logged-in state
    session.clear()
    # Redirect to the index.html page
    return redirect(url_for('main.index'))    

@main.route('/todays-report', methods=['GET'])
def todays_report():
    from datetime import datetime

    # Get today's date
    today = datetime.now().strftime('%B %d, %Y')  # Format as "Month Day, Year"


    try:
        # Load all sheets from the Excel file
        excel_data = pd.read_excel(EXCEL_FILE, sheet_name=None)  # Load all sheets as a dictionary
        sheets_data = {}    

        for sheet_name, df in excel_data.items():
            # Skip the first row of the DataFrame
            df = df.iloc[1:].reset_index(drop=True)

            # Remove the word "Unnamed:" from column names
            df = df.rename(columns=lambda x: str(x).replace('Unnamed:', '').strip())

            # Replace NaN values with empty strings
            df = df.fillna('')

            # Convert numeric columns to integers where possible
            df = df.applymap(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x)

            # Store the cleaned data for each sheet
            sheets_data[sheet_name] = df.to_dict(orient='records')

        # Render the data in the HTML template
        return render_template('todays_report.html', sheets_data=sheets_data, today=today)

    except FileNotFoundError:
        # Handle the case where the Excel file is not found
        return render_template('todays_report.html', sheets_data={}, today=today)