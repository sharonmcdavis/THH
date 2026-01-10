import json

def load_data_from_file(file_path="app_data.json"):
    print("...in load_data_from_file")
    """Load data from a JSON file."""
    try:
        with open(file_path, "r") as file:
            data = json.load(file)
        # print(data)
        data['students'] = sorted(data.get('students', []))
        return data
    except FileNotFoundError:
        return {"students": [], "times": [], "column1_options": {}, "column2_options": {}, "column3_options": {}, "column4_options": {}}

def save_data_to_file(data, file_path="app_data.json"):
    print("...in save_data_to_file")
    """Save data to a JSON file."""
    try:
        with open(file_path, "w") as file:
            json.dump(data, file, indent=4)
        print(f"Data successfully saved to {file_path}")
    except Exception as e:
        print(f"Error saving data: {e}")