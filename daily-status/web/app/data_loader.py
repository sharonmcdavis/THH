import json
from .utils import DATA_FILE

def load_data_from_file():
    """Load data from a JSON file."""
    try:
        with open(DATA_FILE, "r") as file:
            data = json.load(file)
        # print(data)
        # data['students'] = sorted(data.get('students', {}))
        data['students'] = dict(sorted(data.get('students', {}).items()))
        print(data['students'])
        return data
    except FileNotFoundError:
        return {"students": {}, "times": [], "column1_options": {}, "column2_options": {}, "column3_options": {}, "column4_options": {}}

def save_data_to_file(data):
    print("...in save_data_to_file")
    """Save data to a JSON file."""
    try:
        with open(DATA_FILE, "w") as file:
            json.dump(data, file, indent=4)
        print(f"Data successfully saved to {DATA_FILE}")
    except Exception as e:
        print(f"Error saving data: {e}")