# data_loader.py
import os
import json

DATA_FILE = "app_data.json"

def load_data_from_file():
    """Load data from the JSON file."""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as file:
            return json.load(file)
    else:
        return {
            "students": [],
            "times": [],
            "column1": [],
            "column2": [],
            "column3": [],
            "column4": [],
        }

def save_data_to_file(data):
    """Save data to the JSON file."""
    with open(DATA_FILE, "w") as file:
        json.dump(data, file)