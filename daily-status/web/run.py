from app import create_app
from app.data_storage import initialize_data

app = create_app()

# Initialize global data on startup
initialize_data()

if __name__ == '__main__':
    app.run(debug=True)