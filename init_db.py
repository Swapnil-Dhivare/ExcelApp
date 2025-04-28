import os
import sys
from datetime import datetime

# Include the current directory in the path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app_updated import app, db, User, Sheet

def init_db():
    """Initialize the database with tables and sample data"""
    with app.app_context():
        # Create tables
        db.create_all()
        
        # Check if we already have users
        if User.query.count() > 0:
            print("Database already contains users, skipping initialization")
            return

        # Create demo user
        demo_user = User(
            username='demo',
            email='demo@example.com',
            name='Demo User',
            created_at=datetime.utcnow()
        )
        demo_user.set_password('password')
        db.session.add(demo_user)
        
        # Create admin user
        admin_user = User(
            username='admin',
            email='admin@example.com',
            name='Admin User',
            created_at=datetime.utcnow()
        )
        admin_user.set_password('admin123')
        db.session.add(admin_user)
        
        # Commit users first to get their IDs
        db.session.commit()
        
        # Create a sample sheet for the demo user
        sample_sheet = Sheet(
            sheet_name='Sample Sheet',
            data=[
                ['Name', 'Age', 'City', 'Occupation', 'Salary'],
                ['John Doe', '28', 'New York', 'Engineer', '85000'],
                ['Jane Smith', '32', 'San Francisco', 'Designer', '92000'],
                ['Bob Johnson', '45', 'Chicago', 'Manager', '105000'],
                ['Alice Brown', '29', 'Seattle', 'Developer', '95000']
            ],
            user_id=demo_user.id,  # Use the actual user ID
            created_at=datetime.utcnow(),
            updated_at=datetime.utcnow()
        )
        db.session.add(sample_sheet)
        
        # Commit all changes
        db.session.commit()
        print("Database initialized successfully with demo data!")

def reset_db():
    """Delete and recreate all tables"""
    with app.app_context():
        db.drop_all()
        db.create_all()
        print("Database has been reset!")

if __name__ == '__main__':
    if len(sys.argv) > 1:
        if sys.argv[1] == 'reset':
            confirm = input("This will DELETE ALL DATA in the database. Are you sure? (yes/no): ")
            if confirm.lower() == 'yes':
                reset_db()
            else:
                print("Operation cancelled.")
        elif sys.argv[1] == 'init-only':
            with app.app_context():
                db.create_all()
                print("Tables created (no sample data added)")
    else:
        init_db()