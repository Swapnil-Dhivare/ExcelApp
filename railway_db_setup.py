import sys
import os
import pymysql

# Add the current directory to the Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import configuration
from config import DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_NAME

def test_connection():
    """Test connection to Railway MySQL database"""
    try:
        # Connect to the Railway MySQL server
        connection = pymysql.connect(
            host=DB_HOST,
            port=int(DB_PORT),
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME
        )
        
        if connection.open:
            cursor = connection.cursor()
            # Get MySQL version
            cursor.execute("SELECT VERSION()")
            version = cursor.fetchone()
            print(f"Connected to Railway MySQL server successfully")
            print(f"MySQL version: {version[0]}")
            
            # Show existing tables
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()
            
            if tables:
                print("\nExisting tables:")
                for table in tables:
                    print(f" - {table[0]}")
            else:
                print("\nNo tables exist in the database yet.")
                
            cursor.close()
            connection.close()
            print("MySQL connection is closed")
            return True
    except Exception as e:
        print(f"Error while connecting to Railway MySQL: {e}")
        return False

if __name__ == "__main__":
    print("Testing connection to Railway MySQL database...")
    success = test_connection()
    
    if success:
        print("\nConnection test successful!")
        print("\nYou can now run the initialization script to create tables:")
        print("python init_db.py")
    else:
        print("\nConnection test failed!")
        print("Please check your Railway MySQL credentials in config.py")
