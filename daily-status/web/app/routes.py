import os
import shutil
from zoneinfo import ZoneInfo
from flask import Blueprint, json, jsonify, render_template, request, redirect, url_for, flash, session
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

    if (not time):
        # Get the current time in UTC
        utc_now = datetime.now(ZoneInfo("UTC"))

        # Convert UTC time to CST
        cst_now = utc_now.astimezone(ZoneInfo("America/Chicago"))

        # Format the time in 12-hour format with AM/PM
        time = cst_now.strftime('%I:%M')  # Example: "10:05 PM"


        print("Current Time:", time)

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

@main.route('/full-report', methods=['GET'])
def full_report():
    from datetime import datetime

    # Get today's date
    today = datetime.now().strftime('%B %d, %Y')  # Format as "Month Day, Year"
    sheets_data = get_report_data()

    # Render the data in the HTML template
    return render_template('full_report.html', sheets_data=sheets_data, today=today)

    
@main.route('/daily-report', methods=['GET'])
def daily_report():
    print('get daily report')
    try:
        # Get the current date in the format matching your Excel column headers
        display_date = datetime.now().strftime('%B %d, %Y')
        today = datetime.now().strftime('%d')

        # Read the Excel file
        excel_data = pd.ExcelFile(EXCEL_FILE)
        sheets_data = {}

        # Loop through each sheet and filter data for the current day's column
        for sheet_name in excel_data.sheet_names:
            df = excel_data.parse(sheet_name)

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

            # Extract the first column dynamically (starting at row 2)
            time_column = df.iloc[1:, 0].reset_index(drop=True)  # First column, skipping the first row
            if today in df.columns:
                # Extract the data for the current day's column
                today_column = df[today].iloc[1:].reset_index(drop=True)  # Skip the first row for the "today" column

                # Combine the time column and today's column into a list of dictionaries
                combined_data = [
                    {'Time': time, 'Value': value}
                    for time, value in zip(time_column, today_column)
                ]

                # Store the combined data in the sheets_data dictionary
                sheets_data[sheet_name] = combined_data
                print("sheets_data:", sheets_data)

        # Render the data in the new template
        return render_template('daily_report.html', sheets_data=sheets_data, today=display_date)

    except Exception as e:
        # Handle errors gracefully
        return render_template('daily_report.html', today=display_date)
    
def get_report_data():
    try:
        # Read the Excel file
        excel_data = pd.ExcelFile(EXCEL_FILE)
        
        # Initialize a dictionary to store cleaned data for all sheets
        sheets_data = {}

        # Loop through each sheet in the Excel file
        for sheet_name in excel_data.sheet_names:
            print(f"RD - Processing sheet: {sheet_name}")

            # Parse the sheet into a DataFrame
            df = excel_data.parse(sheet_name)

            # Clean the DataFrame
            # Remove the first row if it's a header row or unnecessary
            df = df.iloc[1:].reset_index(drop=True)

            # Strip column names of extra spaces and "Unnamed:" prefixes
            df = df.rename(columns=lambda x: str(x).strip().replace('Unnamed:', ''))

            # Replace NaN values with empty strings
            df = df.fillna('')

            # Convert numeric columns to integers where possible
            df = df.apply(lambda col: col.map(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x))  # Convert numeric columns

            # Store the cleaned data in the sheets_data dictionary
            sheets_data[sheet_name] = df.to_dict(orient='records')
            print(f"RD - Sheet '{sheet_name}' processed. Rows: {len(df)}")

        # Return the cleaned data for all sheets
        return sheets_data

    except Exception as e:
        print(f"An error occurred while processing the Excel file: {e}")
        return {}
            
@main.route('/student-daily-report', methods=['GET'])
def student_daily_report():
    print("get student report")
    try:
        student = request.args.get('student')  # Get the selected student's name from the query parameter
        print("student:", student)

        # Get the current date in the format matching your Excel column headers
        display_date = datetime.now().strftime('%B %d, %Y')
        today = datetime.now().strftime('%d')

        # Read the Excel file
        excel_data = pd.ExcelFile(EXCEL_FILE)
        sheets_data = {}

        # Loop through each sheet and filter data for the current day's column
        for sheet_name in excel_data.sheet_names:
            # print("**parsed: ", sheet_name.split('-')[:-1])
            parsed = sheet_name.split('-')[:-1]

            student_name = '-'.join(parsed).strip()  # Join the name parts with a space and strip whitespace
            print("parsed student: ", student_name)
            df = excel_data.parse(sheet_name)

            if student:
                print("in student: ", df)
                print("student: ", student)
                print("student_name: ", student_name)
                print(f"check: {student.strip().lower() == student_name.lower()}")

                # Compare the student parameter with the first cell (A1)
                if student and student.strip().lower() != student_name.strip().lower():
                    print(f"Skipping sheet '{sheet_name}' as A1 does not match the student parameter.")
                    continue
                else:
                    print("not skipping")

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

            # Extract the first column dynamically (starting at row 2)
            time_column = df.iloc[1:, 0].reset_index(drop=True)  # First column, skipping the first row
            if today in df.columns:
                # Extract the data for the current day's column
                today_column = df[today].iloc[1:].reset_index(drop=True)  # Skip the first row for the "today" column

                # Combine the time column and today's column into a list of dictionaries
                combined_data = [
                    {'Time': time, 'Value': value}
                    for time, value in zip(time_column, today_column)
                ]

                # Store the combined data in the sheets_data dictionary
                sheets_data[sheet_name] = combined_data
                print("sheets_data:", sheets_data)
        return render_template('daily_report.html', sheets_data=sheets_data, today=display_date)

    except Exception as e:
        # flash(f"An error occurred while processing the Excel file: {str(e)}", 500  )
        print(f'An error occurred while processing the Excel file: {e}')
        return render_template('daily_report.html', today=display_date)
    
