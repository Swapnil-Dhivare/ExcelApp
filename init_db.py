from app_updated import app, db
from flask_migrate import Migrate, init, migrate, upgrade

migrate = Migrate(app, db)

with app.app_context():
    # Initialize migrations
    init()
    
    # Create the migration
    migrate()
    
    # Apply the migration
    upgrade()
    
    print("Database initialized successfully!")