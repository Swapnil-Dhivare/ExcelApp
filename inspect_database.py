import sys
import os
import pymysql

# Add the current directory to the Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import configuration
from config import DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_NAME

def inspect_table_structure(table_name):
    """Inspect the structure of a table"""
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
            
            # Get table structure
            cursor.execute(f"DESCRIBE {table_name}")
            columns = cursor.fetchall()
            
            print(f"\nTable structure for '{table_name}':")
            print("-" * 80)
            print(f"{'Field':<20} {'Type':<20} {'Null':<6} {'Key':<6} {'Default':<20} {'Extra':<20}")
            print("-" * 80)
            
            for column in columns:
                print(f"{column[0]:<20} {column[1]:<20} {column[2]:<6} {column[3]:<6} {str(column[4]):<20} {column[5]:<20}")
            
            cursor.close()
            connection.close()
            return True
    except Exception as e:
        print(f"Error inspecting table structure: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        table_name = sys.argv[1]
    else:
        table_name = "user"
    
    inspect_table_structure(table_name)
