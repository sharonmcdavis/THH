import tkinter as tk
from tkinter import messagebox
import json
from admin_window import admin_button_handler
from data_storage import write_to_excel
from data_loader import DATA_FILE

class MainWindow:
    def __init__(self, students, times, column1, column2, column3, column4):
        self.students = students
        self.times = times
        self.column1 = column1
        self.column2 = column2
        self.column3 = column3
        self.column4 = column4
        self.root = tk.Tk()
        self.selected_student = tk.StringVar()
        self.selected_time = tk.StringVar()
        self.notes_text_input = None
        self.column_vars = {
            "Column 1": {key: tk.StringVar(value=value) for key, value in self.column1.items()},
            "Column 2": {key: tk.StringVar(value=value) for key, value in self.column2.items()},
            "Column 3": {key: tk.StringVar(value=value) for key, value in self.column3.items()},
            "Column 4": {key: tk.StringVar(value=value) for key, value in self.column4.items()},
        }

    def create_main_window(self, refresh_callback=None):
        """Create the main application window with scrollable content."""
        self.root.title("Student Activity Tracker")
        self.root.geometry("1200x800")

        # Create a scrollable canvas
        canvas = tk.Canvas(self.root)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        # Configure the canvas and scrollbar
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Place the canvas and scrollbar in the root window
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Create the main layout inside the scrollable frame
        top_frame = tk.Frame(scrollable_frame, name="top_frame")
        top_frame.pack(side="top", fill="x")

        main_frame = tk.Frame(scrollable_frame, name="main_frame")
        main_frame.pack(fill="both", expand=True)

        # Create UI components
        self.create_students(main_frame, self.students, self.selected_student)
        self.create_times(main_frame, self.times, self.selected_time)
        self.create_toggle_buttons(main_frame)
        self.create_submit_buttons(main_frame)

        # Start the Tkinter main loop (only if it hasn't already started)
        if not hasattr(self, "_mainloop_started"):
            self._mainloop_started = True
            self.root.mainloop()

    def create_students(self, parent, students, selected_student):
        students = sorted(students)
        frame = tk.Frame(parent, name="student_frame")
        frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")
        tk.Label(frame, text="Students", font=("Arial", 12, "bold")).grid(row=0, column=0, pady=5)
        for i, student in enumerate(students):
            tk.Radiobutton(
                frame, text=student, variable=selected_student, value=student,
                indicatoron=False, width=15, height=2, bg="#D8BFD8", selectcolor="purple"
            ).grid(row=i + 1, column=0, pady=2)

    def create_times(self, parent, times, selected_time):
        frame = tk.Frame(parent, name="time_frame")
        frame.grid(row=0, column=1, padx=10, pady=10, sticky="n")
        tk.Label(frame, text="Times", font=("Arial", 12, "bold")).grid(row=0, column=0, pady=5)
        for i, time in enumerate(times):
            tk.Radiobutton(
                frame, text=time, variable=selected_time, value=time,
                indicatoron=False, width=15, height=2, bg="lightgreen", selectcolor="green"
            ).grid(row=i + 1, column=0, pady=2)

    # Function to refresh the main window's data
    def refresh_main_window(self):
        """Refresh the main window and reset selections."""
        print("refresh_main_window")

        # Reload data from the JSON file
        try:
            with open(DATA_FILE, "r") as file:
                self.data = json.load(file)  # Reload the updated data
        except Exception as e:
            print(f"Error loading data: {e}")
            return

        print(self.data)
        # Update individual variables with the reloaded data
        self.students = self.data.get("students", [])
        self.times = self.data.get("times", [])
        self.column1 = self.data.get("column1", {})
        self.column2 = self.data.get("column2", {})
        self.column3 = self.data.get("column3", {})
        self.column4 = self.data.get("column4", {})
        
        # Reset the selected student and time variables
        self.selected_student.set("")  # Unselect the student
        self.selected_time.set("")  # Unselect the time

        # Clear all widgets in the main window
        for widget in self.root.winfo_children():
            widget.destroy()

        # Recreate the main window layout with the updated data
        self.create_main_window(refresh_callback=self.refresh_main_window)

    def create_column(self, parent, column_data, column_name, start_row):
        """Create a column section with radio buttons for single selection."""
        tk.Label(parent, text=column_name, font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w")

        # Create a StringVar to track the selected option for this column
        self.column_vars[column_name] = tk.StringVar(value="UNSELECTED")  # Default to no selection

        for idx, (key, value) in enumerate(column_data.items()):
            var_value = tk.StringVar(value="None Selected")
            tk.Radiobutton(
                parent,
                text=value,  # Display the value as the label
                variable=self.column_vars[column_name],  # Bind to the column's StringVar
                value=key,  # The value of this option
                font=("Arial", 10)
            ).grid(sticky="w")

        # Add a horizontal line below the radio buttons
        tk.Frame(parent, height=2, bd=1, relief="sunken").grid(sticky="we", pady=5)

        return start_row + len(column_data) + 2

    def create_toggle_buttons(self, parent):
        """Create the third column with toggle buttons and sections."""
        toggle_frame = tk.Frame(parent, name="options_frame")
        toggle_frame.grid(row=0, column=2, padx=20, pady=10, sticky="n")  # Use grid for toggle_frame

        current_row = 0
        current_row = self.create_column(toggle_frame, self.column1, "Column 1", current_row)
        current_row = self.create_column(toggle_frame, self.column2, "Column 2", current_row)
        current_row = self.create_column(toggle_frame, self.column3, "Column 3", current_row)
        current_row = self.create_column(toggle_frame, self.column4, "Column 4", current_row)

        # Add a text input section
        tk.Label(toggle_frame, text="Notes:", font=("Arial", 10, "bold")).grid(
            row=current_row, column=0, sticky="w", pady=(10, 5)
        )
        self.notes_text_input = tk.Text(toggle_frame, height=5, width=33, wrap="word", font=("Arial", 10))
        self.notes_text_input.grid(row=current_row + 1, column=0, pady=5, sticky="w")
        
    def create_submit_buttons(self, parent):
        """Create the buttons section on the right side of the main window."""
        button_frame = tk.Frame(parent, name="button_frame")
        button_frame.grid(row=0, column=7, rowspan=10, padx=50, pady=20, sticky="ne")

        tk.Button(
            button_frame, text="Save", command=self.submit_data,
            bg="purple", fg="white", height=2, width=20
        ).grid(row=0, column=0, pady=5)

        tk.Button(
            button_frame, text="Admin Functions", command=lambda: admin_button_handler(self.root, self.refresh_main_window),
            bg="green", fg="white", height=2, width=20
        ).grid(row=1, column=0, pady=5)

        tk.Button(
            button_frame, text="Exit Application", command=self.root.destroy,
            bg="gray", fg="white", height=2, width=20
        ).grid(row=2, column=0, pady=5)

    def get_notes_text(self):
        """Retrieve the text from the notes section and replace newlines with spaces."""
        try:
            # Assuming the notes section is a Text widget stored in self.notes_text_input
            text = self.notes_text_input.get("1.0", "end-1c").strip()
            return text.replace("\n", " ")  # Replace newlines with spaces
        except AttributeError:
            messagebox.showerror("Error", "Notes section not found.")
            return ""
        
    def submit_data(self):
        """Handle the save button click."""
        if not self.selected_student.get():
            messagebox.showerror("Error", "Please select a student.")
            return
        if not self.selected_time.get():
            messagebox.showerror("Error", "Please select a time.")
            return

        # Collect selected values for each column
        column_data = {}
        for column_name, var in self.column_vars.items():
            column_data[column_name] = var.get()  # Get the selected value for the column

        # Ensure at least one column has a selection
        if all(value == "UNSELECTED" for value in column_data.values()):
            messagebox.showerror("Error", "Please select at least one option in the columns.")
            return
        
        # Collect selected values for each column
        data = {
            "Student": self.selected_student.get(),
            "Time": self.selected_time.get(),
            **column_data,
            "Notes": self.notes_text_input.get("1.0", "end-1c").strip().replace("\n", " "),
        }

        # Write data to Excel and refresh the main window if successful
        if write_to_excel(self, data):
            messagebox.showinfo("Success", "Data saved successfully!")
            self.refresh_main_window()