# admin_window.py
import json
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog
from data_storage import students, times, save_data, DATA_FILE, add_item, update_listbox
from data_storage import remove_item, open_excel_file, convert_to_pdf

def open_admin_window(root, refresh_callback=None):
    """Open the admin window for managing data."""
    admin_window = tk.Toplevel(root)
    admin_window.title("Admin Functions")
    admin_window.geometry("800x800")

    # Make the admin window stay on top of the main window
    admin_window.transient(root)  # Set the admin window as a child of the main window
    admin_window.grab_set()      # Prevent interaction with the main window
    admin_window.focus_set()     # Focus on the admin window

    # Create a canvas and a scrollbar
    canvas = tk.Canvas(admin_window)
    scrollbar = tk.Scrollbar(admin_window, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    # Configure the canvas
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Use grid for the canvas and scrollbar
    canvas.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    scrollbar.grid(row=0, column=1, sticky="ns", padx=10, pady=10)

    # Make the canvas expandable
    admin_window.grid_rowconfigure(0, weight=1)
    admin_window.grid_columnconfigure(0, weight=1)

    # Load the current data from the JSON file
    try:
        with open(DATA_FILE, "r") as file:
            data = json.load(file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load data: {e}")
        return

    # Ensure all expected columns exist in the data dictionary
    for column in ["column1", "column2", "column3", "column4"]:
        if column not in data:
            data[column] = {}

    # Function to save updated data to the JSON file
    def save_data():
        try:
            with open(DATA_FILE, "w") as file:
                json.dump(data, file, indent=4)
            print("Data successfully saved to app_data.json")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")

    # Section for managing students
    tk.Label(scrollable_frame, text="Manage Students", font=("Arial", 12, "bold")).grid(row=0, column=0, sticky="w", pady=5)
    student_listbox = tk.Listbox(scrollable_frame, height=6, width=30)
    student_listbox.grid(row=1, column=0, sticky="w", pady=(5, 20), padx=20)
    update_listbox(student_listbox, data["students"])  # Populate the listbox with existing students
    student_entry = tk.Entry(scrollable_frame, width=30)
    student_entry.grid(row=2, column=0, sticky="w", pady=(10, 5), padx=20)

    tk.Button(
        scrollable_frame,
        text="Add Student",
        command=lambda: add_item(student_entry, data["students"], student_listbox, save_data),
        bg="green",
        fg="white",
        width=15,
        height=2
    ).grid(row=2, column=1, sticky="n", padx=15)

    tk.Button(
        scrollable_frame,
        text="Remove Student",
        command=lambda: remove_item(student_listbox, data["students"], save_data),
        bg="purple",
        fg="white",
        width=15,
        height=2
    ).grid(row=1, column=1, sticky="n", padx=15)

    # Section for managing times
    tk.Label(scrollable_frame, text="Manage Times", font=("Arial", 12, "bold")).grid(row=3, column=0, sticky="w", pady=5)
    time_listbox = tk.Listbox(scrollable_frame, selectmode="single", height=10, width=30)
    time_listbox.grid(row=4, column=0, sticky="w", pady=(5, 20), padx=20)
    update_listbox(time_listbox, data["times"])  # Populate the listbox with existing times
    time_entry = tk.Entry(scrollable_frame, width=30)
    time_entry.grid(row=5, column=0, sticky="w", pady=(10, 5), padx=20)

    tk.Button(
        scrollable_frame,
        text="Add Time",
        command=lambda: add_item(time_entry, data["times"], time_listbox, save_data),
        bg="green",
        fg="white",
        width=15,
        height=2
    ).grid(row=5, column=1, sticky="n", padx=15)

    tk.Button(
        scrollable_frame,
        text="Remove Time",
        command=lambda: remove_item(time_listbox, data["times"], save_data),
        bg="purple",
        fg="white",
        width=15,
        height=2
    ).grid(row=4, column=1, sticky="n", padx=15)

    # Section for managing column options
    # Dynamically create Listboxes and buttons for each column
    row_offset = 6
    for i, column_name in enumerate(["column1", "column2", "column3", "column4"]):
        tk.Label(scrollable_frame, text=f"Manage {column_name.capitalize()}", font=("Arial", 12, "bold")).grid(
            row=row_offset + i * 3, column=0, sticky="w", pady=5
        )

        # Create a Listbox for the column
        listbox = tk.Listbox(scrollable_frame, height=6, width=30)
        listbox.grid(row=row_offset + i * 3 + 1, column=0, sticky="w", pady=(5, 20), padx=20)

        # Populate the Listbox with existing options
        for key, value in data[column_name].items():
            listbox.insert("end", f"{key}: {value}")

        # Create unique key_entry and value_entry for each column
        key_entry = tk.Entry(scrollable_frame, width=15)
        key_entry.grid(row=row_offset + i * 3 + 2, column=0, sticky="w", pady=(10, 5), padx=20)
        key_entry.insert(0, "Key")

        value_entry = tk.Entry(scrollable_frame, width=30)
        value_entry.grid(row=row_offset + i * 3 + 2, column=1, sticky="w", pady=(10, 5), padx=20)
        value_entry.insert(0, "Value")

        # Add a button to remove an item
        tk.Button(
            scrollable_frame,
            text="Remove Option",
            command=lambda col=column_name, lb=listbox: remove_column_item(col, lb),
            bg="purple",
            fg="white",
            width=15,
            height=2
        ).grid(row=row_offset + i * 3 + 1, column=2, sticky="n", padx=15)

        # Add a button to add an item
        tk.Button(
            scrollable_frame,
            text="Add Option",
            command=lambda col=column_name, ke=key_entry, ve=value_entry, lb=listbox: add_column_item(col, ke, ve, lb, data, save_data),
            bg="green",
            fg="white",
            width=15,
            height=2
        ).grid(row=row_offset + i * 3 + 2, column=2, sticky="n", padx=15)


        # Define the add_column_item function
        def add_column_item(column_name, key_entry, value_entry, listbox, data, save_data):
            """Add a new key-value pair to the specified column."""
            key = key_entry.get().strip()
            value = value_entry.get().strip()

            print("key:", key)
            print("value:", value)
            print("Listbox widget:", listbox)
            print("column_name:", column_name)

            if not key or not value:
                messagebox.showerror("Error", "Key and Value cannot be empty.")
                return

            # Check if the key already exists
            if key in data[column_name]:
                messagebox.showerror("Error", f"Key '{key}' already exists in {column_name}.")
                return

            # Add the new key-value pair to the data dictionary
            data[column_name][key] = value
            save_data()  # Save the updated data to the JSON file

            # Update the Listbox
            listbox.insert("end", f"{key}: {value}")
            print(f"Added '{key}: {value}' to column '{column_name}'.")

            # Clear the key and value entry boxes
            key_entry.delete(0, "end")
            value_entry.delete(0, "end")

        # Function to remove a column item
        def remove_column_item(column, listbox):
            print("remove_column_item")
            selected = listbox.curselection()  # Get the index of the selected item
            print("Selected index:", selected)
            print("Listbox contents:", listbox.get(0, "end"))
            print("Listbox widget:", listbox)

            if not selected:
                messagebox.showerror("Error", "No item selected to remove.")
                return

            # Get the key of the selected item
            selected_index = selected[0]
            selected_item = listbox.get(selected_index)  # Get the selected item as "key: value"
            print("Selected item:", selected_item)
            key = selected_item.split(":")[0].strip()  # Extract the key from "key: value"

            # Remove the item from the data dictionary
            if key in data[column]:
                del data[column][key]
                save_data()  # Save the updated data to the JSON file

                # Update the Listbox
                listbox.delete(selected_index)
                print(f"Item '{key}' removed from column '{column}'.")
            else:
                messagebox.showerror("Error", f"Key '{key}' does not exist in {column}.")



            # tk.Button(
            #     scrollable_frame,
            #     text="Add Option",
            #     command=add_column_item,
            #     bg="green",
            #     fg="white",
            #     width=15,
            #     height=2
            # ).grid(row=row_offset + i * 3 + 2, column=2, sticky="n", padx=15)

            # print("index:")
            # print(i)
            # print("column_name: " + column_name)

            
            # tk.Button(
            #     scrollable_frame,
            #     text="Remove Option",
            #     command=lambda: remove_column_item(column_name, listbox, data, save_data),
            #     bg="purple",
            #     fg="white",
            #     width=15,
            #     height=2
            # ).grid(row=row_offset + i * 3 + 1, column=2, sticky="n", padx=15)

        # Submit, View Excel, and View PDF buttons (right column)
        button_frame = tk.Frame(scrollable_frame)
        button_frame.grid(row=0, column=2, rowspan=10, padx=50, pady=10, sticky="n")  # Align to the top-right

        # Exit Button
        tk.Button(
            button_frame,
            text="Exit",
            command=lambda: (admin_window.destroy(), refresh_callback() if refresh_callback else None),
            bg="gray",
            fg="white",
            width=15,
            height=2
        ).grid(row=0, column=0, sticky="n", columnspan=2, pady=5)

        # Excel Button
        tk.Button(
            button_frame,
            text="View Excel File",
            command=lambda: open_excel_file(),  # Add parentheses to call the function
            bg="gray",
            fg="white",
            width=15,
            height=2
        ).grid(row=1, column=0, sticky="n", columnspan=2, pady=5)

        # PDF Button
        tk.Button(
            button_frame,
            text="View PDF File",
            command=lambda: convert_to_pdf(),  # Add parentheses to call the function
            bg="gray",
            fg="white",
            width=15,
            height=2
        ).grid(row=2, column=0, sticky="n", columnspan=2, pady=5)

def admin_button_handler(root, refresh_callback):
    """Handle the Admin Functions button click with a password prompt."""
    password = simpledialog.askstring("Password Required", "Enter the admin password:", show="*")
    if password == "1102":
        open_admin_window(root, refresh_callback)
    else:
        messagebox.showerror("Access Denied", "Incorrect password. Access to admin functions is denied.")

def save_and_close(admin_window):
    """Save data and close the admin window."""
    # Add logic to save data (e.g., call save_data from data_storage)
    print("Data saved!")
    save_data()
    admin_window.destroy()  # Close the admin window

def load_data():
    global students, times, column1, column2, column3, column4
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as file:
            data = json.load(file)
            students = data.get("students", [])
            times = data.get("times", [])
            column1 = data.get("column1", [])
            column2 = data.get("column2", [])
            column3 = data.get("column3", [])
            column4 = data.get("column4", [])
    else:
        # Initialize with default values if the file does not exist
        students = []
        times = []
        column1 = []
        column2 = []
        column3 = []
        column4 = []

# After calling load_data()
load_data()

# print("in admin_window.py:")
# print("Students:", students)
# print("Times:", times)

