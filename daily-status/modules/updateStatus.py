import data_storage  # Import the module instead of the variables
from main_window import MainWindow

# Initialize data
data_storage.initialize_data()

# Create and run the main application window
main_window = MainWindow(
    data_storage.students,
    data_storage.times,
    data_storage.column1,
    data_storage.column2,
    data_storage.column3,
    data_storage.column4
)

main_window.create_main_window(refresh_callback=lambda: None)  # Replace `None` with the actual refresh logic if needed