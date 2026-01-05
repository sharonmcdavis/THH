# data_utils.py
from tkinter import messagebox

# def update_listbox(listbox, items):
#     """Update a Tkinter listbox with the given items."""
#     listbox.delete(0, "end")  # Clear the listbox
#     for item in items:
#         listbox.insert("end", item)  # Add each item to the listbox

# def add_item(entry_widget, target_list, listbox, save_callback):
#     """Add an item to the list with validation."""
#     item = entry_widget.get().strip()
#     if not item:
#         messagebox.showerror("Error", "Item cannot be empty.")
#         return
#     if item in target_list:
#         messagebox.showerror("Error", f"'{item}' already exists.")
#         return
#     target_list.append(item)  # Add the item to the list
#     update_listbox(listbox, target_list)  # Update the listbox
#     save_callback()  # Save the updated data to the JSON file
#     messagebox.showinfo("Success", f"'{item}' added successfully!")
#     entry_widget.delete(0, "end")  # Clear the entry widget

# def remove_item(listbox, target_list, save_callback):
#     """Remove an item from the list with validation."""
#     selected_index = listbox.curselection()
#     if not selected_index:
#         messagebox.showerror("Error", "No item selected.")
#         return
#     item = listbox.get(selected_index)
#     target_list.remove(item)  # Remove the item from the list
#     update_listbox(listbox, target_list)  # Update the listbox
#     save_callback()  # Save the updated data to the JSON file
#     messagebox.showinfo("Success", f"'{item}' removed successfully!")

