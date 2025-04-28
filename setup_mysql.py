import mysql.connector
from mysql.connector import Error
import os
from config import DB_USER, DB_PASSWORD, DB_HOST, DB_NAME

def create_database():
    """Create MySQL database and user if they don't exist"""
    try:
        # Connect to MySQL server
        connection = mysql.connector.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD
        )
        
        if connection.is_connected():
            cursor = connection.cursor()
            
            # Create database if it doesn't exist
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_NAME}")
            print(f"Database '{DB_NAME}' created or already exists")
            
            cursor.close()
            connection.close()
            print("MySQL connection is closed")
            
            return True
    except Error as e:
        print(f"Error while connecting to MySQL: {e}")
        return False

if __name__ == "__main__":
    if create_database():
        print("Database setup completed successfully")
        print("\nNow you can run the following command to initialize the tables:")
        print("python init_db.py")
    else:
        print("Database setup failed")
        print("\nPlease check your MySQL connection settings in config.py")
