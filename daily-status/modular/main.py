import tkinter as tk
from data_manager import load_data
from main_window import create_main_window

# Load data from the file
students, times, column1, column2, column3, column4 = load_data()

# Create the main application window
root = tk.Tk()
root.title("Student Activity Tracker")
root.geometry("1200x800")  # Set the window size

# Create the main window
create_main_window(root, students, times, column1, column2, column3, column4)

# Run the application
root.mainloop()