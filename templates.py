"""
Templates module for Excel Generator.
This module provides built-in templates for creating sheets.
"""

class SheetTemplate:
    """Represents a template for creating new sheets"""
    
    def __init__(self, id, name, headers=None, sample_data=None, default_sheet_name=None, description=None):
        """
        Initialize a new sheet template
        
        Args:
            id: Unique identifier for the template
            name: Display name of the template
            headers: List of column headers
            sample_data: List of lists representing rows and cells
            default_sheet_name: Default name for sheets created from this template
            description: Description of the template
        """
        self.id = id
        self.name = name
        self.headers = headers or []
        self.sample_data = sample_data or []
        self.default_sheet_name = default_sheet_name or name
        self.description = description or ""
        self.sample_data_string = ""  # Will be computed when needed

    def __str__(self):
        return f"Template: {self.name} ({len(self.headers)} columns)"

def get_built_in_templates():
    """Return a list of built-in sheet templates"""
    templates = [
        # Assignment Tracker
        SheetTemplate(
            id="1",
            name="Assignment Tracker",
            default_sheet_name="Assignment Tracker",
            headers=["Sr No", "Name", "Roll No", "ASS1", "ASS2", "ASS3", "Total"],
            sample_data=[
                ["1", "John Doe", "101", "85", "90", "88", "=SUM(D2:F2)"],
                ["2", "Jane Smith", "102", "78", "92", "95", "=SUM(D3:F3)"],
                ["3", "Bob Johnson", "103", "92", "85", "79", "=SUM(D4:F4)"],
                ["4", "Alice Brown", "104", "88", "91", "94", "=SUM(D5:F5)"],
                ["5", "Charlie Wilson", "105", "76", "82", "89", "=SUM(D6:F6)"]
            ],
            description="Track student assignments with auto-calculated totals"
        ),
        
        # Expense Tracker
        SheetTemplate(
            id="2",
            name="Expense Tracker",
            default_sheet_name="Expense Tracker",
            headers=["Date", "Category", "Description", "Amount", "Payment Method", "Notes"],
            sample_data=[
                ["2023-07-01", "Groceries", "Weekly shopping", "78.50", "Credit Card", ""],
                ["2023-07-03", "Utilities", "Electricity bill", "120.00", "Bank Transfer", ""],
                ["2023-07-05", "Transportation", "Gas", "45.75", "Cash", ""],
                ["2023-07-08", "Entertainment", "Movie tickets", "24.00", "Credit Card", ""],
                ["2023-07-10", "Food", "Restaurant dinner", "65.25", "Credit Card", ""]
            ],
            description="Track your daily expenses by category"
        ),
        
        # Project Task List
        SheetTemplate(
            id="3",
            name="Project Task List",
            default_sheet_name="Project Tasks",
            headers=["Task ID", "Task Name", "Assigned To", "Start Date", "Due Date", "Status", "Priority", "Notes"],
            sample_data=[
                ["T-001", "Project Planning", "John Smith", "2023-07-01", "2023-07-05", "Completed", "High", ""],
                ["T-002", "Research", "Jane Doe", "2023-07-06", "2023-07-15", "In Progress", "Medium", ""],
                ["T-003", "Design", "Bob Johnson", "2023-07-16", "2023-07-25", "Not Started", "High", ""],
                ["T-004", "Development", "Alice Brown", "2023-07-26", "2023-08-15", "Not Started", "High", ""],
                ["T-005", "Testing", "Charlie Wilson", "2023-08-16", "2023-08-25", "Not Started", "Medium", ""]
            ],
            description="Manage project tasks with priorities and deadlines"
        ),
        
        # Inventory Tracker
        SheetTemplate(
            id="4",
            name="Inventory Tracker",
            default_sheet_name="Inventory",
            headers=["Item ID", "Item Name", "Category", "Quantity", "Unit Price", "Total Value", "Reorder Level", "Last Updated"],
            sample_data=[
                ["I-001", "Desk Chair", "Furniture", "15", "89.99", "=D2*E2", "5", "2023-06-15"],
                ["I-002", "Desk Lamp", "Lighting", "25", "24.99", "=D3*E3", "10", "2023-06-18"],
                ["I-003", "Notebook", "Stationery", "100", "3.99", "=D4*E4", "25", "2023-06-20"],
                ["I-004", "Ballpoint Pen", "Stationery", "200", "1.49", "=D5*E5", "50", "2023-06-22"],
                ["I-005", "Wireless Mouse", "Electronics", "20", "19.99", "=D6*E6", "8", "2023-06-25"]
            ],
            description="Track inventory with automatic value calculation"
        ),
        
        # Budget Planner
        SheetTemplate(
            id="5",
            name="Budget Planner",
            default_sheet_name="Budget",
            headers=["Category", "Budgeted", "Week 1", "Week 2", "Week 3", "Week 4", "Total Spent", "Remaining"],
            sample_data=[
                ["Groceries", "400.00", "95.25", "105.50", "85.75", "98.30", "=SUM(C2:F2)", "=B2-G2"],
                ["Utilities", "200.00", "0.00", "180.00", "0.00", "0.00", "=SUM(C3:F3)", "=B3-G3"],
                ["Transportation", "150.00", "35.00", "42.50", "38.25", "45.75", "=SUM(C4:F4)", "=B4-G4"],
                ["Entertainment", "100.00", "0.00", "45.00", "25.00", "35.00", "=SUM(C5:F5)", "=B5-G5"],
                ["Dining Out", "120.00", "35.00", "0.00", "42.50", "38.00", "=SUM(C6:F6)", "=B6-G6"]
            ],
            description="Plan and track your monthly budget by category"
        ),
        
        # Employee Attendance
        SheetTemplate(
            id="6",
            name="Employee Attendance",
            default_sheet_name="Attendance",
            headers=["Emp ID", "Name", "Mon", "Tue", "Wed", "Thu", "Fri", "Total Hours", "Overtime"],
            sample_data=[
                ["E001", "John Smith", "8", "8", "8", "8", "8", "=SUM(C2:G2)", "=IF(H2>40,H2-40,0)"],
                ["E002", "Jane Doe", "8", "8", "9", "8", "7", "=SUM(C3:G3)", "=IF(H3>40,H3-40,0)"],
                ["E003", "Bob Johnson", "9", "8", "8", "9", "8", "=SUM(C4:G4)", "=IF(H4>40,H4-40,0)"],
                ["E004", "Alice Brown", "8", "8", "8", "8", "4", "=SUM(C5:G5)", "=IF(H5>40,H5-40,0)"],
                ["E005", "Charlie Wilson", "10", "8", "9", "8", "8", "=SUM(C6:G6)", "=IF(H6>40,H6-40,0)"]
            ],
            description="Track employee attendance and calculate overtime"
        )
    ]
    
    # Pre-compute sample_data_string for each template
    for template in templates:
        if template.sample_data:
            rows = []
            for row in template.sample_data:
                rows.append(','.join(map(str, row)))
            template.sample_data_string = ';'.join(rows)
    
    return templates
