import json
import os

# File to store the lists
data_file = "app_data.json"

# Function to save the lists to a file
def save_data(students, times, column1, column2, column3, column4):
    data = {
        "students": students,
        "times": times,
        "column1": column1,
        "column2": column2,
        "column3": column3,
        "column4": column4,
    }
    with open(data_file, "w") as file:
        json.dump(data, file)

# Function to load the lists from a file
def load_data():
    if os.path.exists(data_file):
        with open(data_file, "r") as file:
            data = json.load(file)
            return (
                data.get("students", []),
                data.get("times", []),
                data.get("column1", []),
                data.get("column2", []),
                data.get("column3", []),
                data.get("column4", []),
            )
    else:
        # Initialize with default values if the file does not exist
        return [], [], [], [], [], []