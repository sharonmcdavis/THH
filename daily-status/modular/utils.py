# Function to handle toggle button clicks
def toggle_button(var, value):
    if var.get() == value:
        var.set("")  # Deselect if clicked again
    else:
        var.set(value)

# Function to reset all buttons
def reset_buttons():
    selected_student.set("")
    selected_time.set("")
    for var in column1_values.values():
        var.set(False)
    for var in column2_values.values():
        var.set(False)
    for var in column3_values.values():
        var.set(False)
    for var in column4_values.values():
        var.set(False)

# Helper function to update the listbox with current items
def update_listbox(listbox, items):
    listbox.delete(0, tk.END)  # Clear the listbox
    for item in items:
        listbox.insert(tk.END, item)  # Add each item to the listbox

# Helper function to add an item to a list and update the listbox
def add_item(entry_widget, target_list, listbox):
    item = entry_widget.get().strip()
    if item and item not in target_list:
        target_list.append(item)
        update_listbox(listbox, target_list)  # Update the listbox
        save_data()  # Save the updated data
        messagebox.showinfo("Success", f"'{item}' added successfully!")
        entry_widget.delete(0, tk.END)
    elif item in target_list:
        messagebox.showerror("Error", f"'{item}' already exists!")
    else:
        messagebox.showerror("Error", "Please enter a valid item.")

# Helper function to modify an item in a list and update the listbox
def modify_item(entry_widget, target_list, listbox):
    selected_index = listbox.curselection()
    if selected_index:
        old_item = listbox.get(selected_index)
        new_item = entry_widget.get().strip()
        if new_item and new_item not in target_list:
            target_list[target_list.index(old_item)] = new_item
            update_listbox(listbox, target_list)  # Update the listbox
            save_data()  # Save the updated data
            messagebox.showinfo("Success", f"'{old_item}' modified to '{new_item}' successfully!")
            entry_widget.delete(0, tk.END)
        elif new_item in target_list:
            messagebox.showerror("Error", f"'{new_item}' already exists!")
        else:
            messagebox.showerror("Error", "Please enter a valid item.")
    else:
        messagebox.showerror("Error", "Please select an item to modify.")

# Helper function to remove an item from a list and update the listbox
def remove_item(entry_widget, target_list, listbox):
    selected_index = listbox.curselection()
    if selected_index:
        item = listbox.get(selected_index)
        target_list.remove(item)
        update_listbox(listbox, target_list)  # Update the listbox
        save_data()  # Save the updated data
        messagebox.showinfo("Success", f"'{item}' removed successfully!")
        entry_widget.delete(0, tk.END)
    else:
        messagebox.showerror("Error", "Please select an item to remove.")

# Helper function to get the appropriate column list
def get_column_list(column_name):
    if column_name == "Column1":
        return column1
    elif column_name == "Column2":
        return column2
    elif column_name == "Column3":
        return column3
    elif column_name == "Column4":
        return column4        