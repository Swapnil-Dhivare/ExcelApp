import sys
import os
import pymysql

# Add the current directory to the Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import configuration
from config import DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_NAME

def execute_sql(sql):
    """Execute SQL on the Railway MySQL database"""
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
            print(f"Executing SQL: {sql}")
            cursor.execute(sql)
            connection.commit()
            print("SQL executed successfully")
            cursor.close()
            connection.close()
            return True
    except Exception as e:
        print(f"Error executing SQL: {e}")
        return False

def alter_password_hash_column():
    """Alter the password_hash column in the user table"""
    # SQL to alter the column
    sql = "ALTER TABLE user MODIFY COLUMN password_hash VARCHAR(255)"
    return execute_sql(sql)

if __name__ == "__main__":
    print("Altering password_hash column in user table...")
    if alter_password_hash_column():
        print("Column altered successfully!")
    else:
        print("Failed to alter column.")
