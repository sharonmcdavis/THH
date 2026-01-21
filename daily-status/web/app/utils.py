import os
from flask import session, redirect, url_for, flash
from datetime import datetime, timedelta

# Define the base directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Define file paths
DATA_FILE = os.path.join(BASE_DIR, "app_data.json")
EXCEL_FILE = os.path.join(BASE_DIR, "student_activity.xlsx")
BACKUP_FOLDER = os.path.join("./backups/")
ARCHIVE_FOLDER = os.path.join("./archive/")

WEB_PASSWORD = "401"
ADMIN_PASSWORD = "1102"

# Middleware to validate login state
def login_required(f):
    def wrapper(*args, **kwargs):
        # Check if the user is logged in
        if not session.get('logged_in'):
            flash("You must log in to access this page.", "error")
            return redirect(url_for('main.index'))
        
        # Check if the session has expired
        login_time = session.get('login_time')
        if login_time:
            login_time = datetime.strptime(login_time, '%Y-%m-%d %H:%M:%S')
            if datetime.now() - login_time > timedelta(minutes=600):
                session.clear()  # Clear the session
                flash("Your session has expired. Please log in again.", "error")
                return redirect(url_for('main.index'))
        session['login_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__  # Preserve the original function name
    return wrapper

def admin_login_required(f):
    def wrapper(*args, **kwargs):
        # Check if the user is logged in
        if not session.get('admin_logged_in'):
            flash("You must log in as Admin to access this page.", "error")
            return redirect(url_for('main.main_window'))

        session['login_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__  # Preserve the original function name
    return wrapper