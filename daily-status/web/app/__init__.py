from flask import Flask

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'THH_secret_key'

    # Register blueprints
    from .routes import main
    from .admin_routes import admin
    app.register_blueprint(main)
    app.register_blueprint(admin, url_prefix='/admin')

    return app