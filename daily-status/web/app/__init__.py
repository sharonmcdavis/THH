from flask import Flask
import colorsys

def lighten(color, factor):
    """Lighten a hex color by a given factor (0.0 to 1.0)."""
    if not color.startswith('#') or len(color) != 7:
        raise ValueError(f"Invalid hex color: {color}")
    color = color.lstrip('#')
    r, g, b = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
    r, g, b = [min(255, int(c + (255 - c) * factor)) for c in (r, g, b)]
    return f'#{r:02x}{g:02x}{b:02x}'

def darken(color, factor):
    """Darken a hex color by a given factor (0.0 to 1.0)."""
    if not color.startswith('#') or len(color) != 7:
        raise ValueError(f"Invalid hex color: {color}")
    color = color.lstrip('#')
    r, g, b = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
    r, g, b = [max(0, int(c * (1 - factor))) for c in (r, g, b)]
    return f'#{r:02x}{g:02x}{b:02x}'

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'THH_secret_key'

    # Register blueprints
    from .routes import main
    from .admin_routes import admin
    app.register_blueprint(main)
    app.register_blueprint(admin, url_prefix='/admin')

    app.jinja_env.filters['lighten'] = lighten
    app.jinja_env.filters['darken'] = darken

    return app
