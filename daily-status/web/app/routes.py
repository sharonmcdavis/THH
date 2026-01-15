from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from datetime import datetime, timedelta
from .data_storage import initialize_data, write_to_excel
from .utils import login_required, WEB_PASSWORD

main = Blueprint('main', __name__)

@main.route('/', methods=['GET', 'POST'])
def index():
    print("in the main default /")
    if request.method == 'POST':
        username = request.form.get('username')
        print("username:", username)
        password = request.form.get('password')
        if username and password == WEB_PASSWORD:
            # Save login state and timestamp in the session
            session['username'] = username
            print("username saved to session")
            session['logged_in'] = True
            session['login_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            return redirect(url_for('main.main_window'))
        else:
            flash("Access restricted. Invalid login credentials.", "error")
            return render_template('index.html')
    return render_template('index.html')

@main.route('/main')
@login_required
def main_window():
    print("in the main main")

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