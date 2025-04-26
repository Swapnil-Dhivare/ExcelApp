from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin, LoginManager, login_user, logout_user, current_user, login_required
from flask import Flask, render_template, redirect, url_for, request, flash, send_file, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from flask_migrate import Migrate
from datetime import datetime
import os
import json
import io
import xlsxwriter
import pandas as pd
import base64
import logging
from dotenv import load_dotenv

db = SQLAlchemy()

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    sheets = db.relationship('Sheet', backref='user', lazy=True)
    downloads = db.relationship('DownloadHistory', backref='user', lazy=True)

class Sheet(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    sheet_name = db.Column(db.String(100), nullable=False)
    data = db.Column(db.JSON, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class DownloadHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    sheet_data = db.Column(db.JSON)
    downloaded_at = db.Column(db.DateTime, default=datetime.utcnow)

# Load environment variables
load_dotenv()

# Initialize app
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///app.db'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.logger.setLevel(logging.DEBUG)

# Configuration
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Initialize login manager
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Initialize database
db.init_app(app)
migrate = Migrate(app, db)

def init_db():
    with app.app_context():
        # Drop all tables
        db.drop_all()
        # Create all tables
        db.create_all()

with app.app_context():
    db.create_all()

# Add this after your app initialization
@app.context_processor
def inject_templates():
    return {
        'built_in_templates': built_in_templates,
        'custom_templates': custom_templates
    }

# Global data stores
sheets_data = []
custom_templates = []
sheet_order = []

# Built-in templates
built_in_templates = [
    {
        "id": "front_page",
        "name": "Front Page",
        "default_sheet_name": "Front Page",
        "headers": ["Title", "Subtitle", "Date", "Author"],
        "sample_data": [
            ["Course Workbook", "Comprehensive Guide", "2023-09-01", "John Doe"],
            ["", "Module 1", "", ""],
            ["", "Module 2", "", ""]
        ],
        "formatting": {
            "title_font_size": 20,
            "title_bold": True,
            "header_bg_color": "#4F81BD",
            "cell_font_size": 12,
            "cell_alignment": "center",
            "cell_border_color": "#000000",
            "text_wrap": True,
            "merge_cells": "A1:D1,A2:D2,A3:D3"
        }
    },
    {
        "id": "summary",
        "name": "Summary Sheet",
        "default_sheet_name": "Summary",
        "headers": ["Item", "Value", "Formula", "Notes"],
        "sample_data": [
            ["Total Students", "=COUNT(Data!B2:B100)", "", "Automatic count"],
            ["Average Score", "=AVERAGE(Data!C2:C100)", "", "Calculated field"]
        ],
        "formatting": {
            "header_bg_color": "#9BBB59",
            "title_font_size": 14,
            "cell_alignment": "left",
            "number_format": "#,##0.00",
            "validation": {"Notes": {"type": "list", "options": ["Important", "Review", "Done"]}}
        }
    },
    {
        'id': 'course_file',
        'name': 'Course File Template',
        'headers': ['Subject', 'Code', 'Class', 'Semester', 'Teacher'],
        'sample_data': [
            ['Distributed Computing', 'CSC801', 'BE', 'VIII', 'Dr. K. T. Patil']
        ]
    },
    {
        'id': 'attendance',
        'name': 'Attendance Sheet',
        'headers': ['Roll No', 'Name', 'Date 1', 'Date 2', 'Date 3', 'Percentage'],
        'sample_data': [
            ['1', 'Student 1', 'P', 'P', 'A', '66.67%'],
            ['2', 'Student 2', 'P', 'P', 'P', '100%']
        ]
    },
    {
        'id': 'marks',
        'name': 'Marks Sheet',
        'headers': ['Roll No', 'Name', 'Quiz 1', 'Quiz 2', 'Final', 'Total'],
        'sample_data': [
            ['1', 'Student 1', '8', '9', '85', '102'],
            ['2', 'Student 2', '9', '10', '88', '107']
        ]
    }
]

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def get_column_letter_index(column_letter):
    """Convert Excel column letter to index"""
    result = 0
    for char in column_letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def get_cell_format(workbook, formatting, row, col):
    """Create cell format based on stored formatting"""
    cell_key = f"{row + 1},{col + 1}"
    cell_format = workbook.add_format()
    
    if cell_key in formatting.get('formatted_cells', {}):
        fmt = formatting['formatted_cells'][cell_key]
        
        # Apply font formatting
        if 'font' in fmt:
            font = fmt['font']
            if font:
                cell_format.set_font_name(font.get('name', 'Calibri'))
                cell_format.set_font_size(float(font.get('size', 11)))
                cell_format.set_bold(font.get('bold', False))
                cell_format.set_italic(font.get('italic', False))
                
                # Handle font color
                if 'color' in font and font['color']:
                    color = font['color'].get('rgb', '')
                    if color:
                        cell_format.set_font_color(color)
        
        # Apply fill
        if 'fill' in fmt:
            fill = fmt['fill']
            if fill and 'fgColor' in fill:
                color = fill['fgColor'].get('rgb', '')
                if color:
                    cell_format.set_pattern(1)
                    cell_format.set_bg_color(color)
        
        # Apply alignment
        if 'alignment' in fmt:
            alignment = fmt['alignment']
            if alignment:
                cell_format.set_align(alignment.get('horizontal', 'left'))
                cell_format.set_align(alignment.get('vertical', 'top'))
        
        # Apply border
        if 'border' in fmt:
            border = fmt['border']
            if border:
                for side in ['top', 'right', 'bottom', 'left']:
                    if side in border:
                        style = border[side].get('style', 'thin')
                        color = border[side].get('color', {}).get('rgb', '#000000')
                        cell_format.set_border(get_border_style(style))
                        cell_format.set_border_color(color)
        
        # Apply number format
        if 'number_format' in fmt:
            cell_format.set_num_format(fmt['number_format'])
            
    return cell_format

def get_border_style(style_name):
    """Convert border style names to xlsxwriter values"""
    style_map = {
        'thin': 1,
        'medium': 2,
        'thick': 3,
        'dotted': 4,
        'dashed': 5,
        'double': 6,
    }
    return style_map.get(style_name, 1)

def serialize_style_object(obj):
    """Convert style objects to serializable dictionaries"""
    if obj is None:
        return None
    
    if isinstance(obj, bytes):
        return serialize_binary_data(obj)
    
    if hasattr(obj, '__dict__'):
        # Convert object attributes to dictionary
        result = {}
        for key, value in obj.__dict__.items():
            if not key.startswith('_'):  # Skip private attributes
                if isinstance(value, bytes):
                    result[key] = serialize_binary_data(value)
                elif hasattr(value, '__dict__'):
                    result[key] = serialize_style_object(value)
                else:
                    result[key] = str(value) if not isinstance(value, (bool, int, float)) else value
        return result
    return str(obj)

def serialize_binary_data(data):
    """Convert binary data to base64 string"""
    if isinstance(data, bytes):
        return base64.b64encode(data).decode('utf-8')
    return data

def serialize_color(color_obj):
    """Convert openpyxl color objects to hex string"""
    if color_obj is None:
        return None
    if hasattr(color_obj, 'rgb'):
        if isinstance(color_obj.rgb, str):
            return color_obj.rgb
        elif isinstance(color_obj.rgb, bytes):
            return color_obj.rgb.hex()
    return None

def serialize_cell_format(cell):
    """Convert cell formatting to serializable dictionary"""
    return {
        'font': {
            'name': cell.font.name or 'Arial',
            'size': cell.font.size or 11,
            'bold': cell.font.bold,
            'italic': cell.font.italic,
            'color': serialize_color(cell.font.color)
        },
        'alignment': {
            'horizontal': cell.alignment.horizontal or 'left',
            'vertical': cell.alignment.vertical or 'center',
            'wrap_text': cell.alignment.wrap_text
        },
        'border': {
            'style': cell.border.outline_level if cell.border else 1,
            'double': any(getattr(cell.border, side).style == 'double' 
                         for side in ['top', 'right', 'bottom', 'left'] 
                         if getattr(cell.border, side))
        },
        'fill': {
            'type': cell.fill.fill_type if cell.fill else None,
            'color': serialize_color(cell.fill.fgColor) if cell.fill else None
        }
    }

@app.route('/')
def root():
    if not current_user.is_authenticated:
        return redirect(url_for('login'))
    return redirect(url_for('index'))

@app.route('/index')
@login_required
def index():
    return render_template(
        'index.html',
        sheets=current_user.sheets,
        built_in_templates=built_in_templates,
        custom_templates=custom_templates
    )

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        remember = request.form.get('remember') == 'on'
        
        user = User.query.filter_by(email=email).first()
        
        if user and check_password_hash(user.password, password):
            login_user(user, remember=remember)  # Set remember=True to maintain login
            next_page = request.args.get('next')
            return redirect(next_page) if next_page else redirect(url_for('index'))
        
        flash('Invalid email or password', 'danger')
    
    return render_template('login.html', 
                        built_in_templates=built_in_templates, 
                        custom_templates=custom_templates)

@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '').strip()
        
        # Validate input
        if not all([name, email, password]):
            flash('All fields are required')
            return redirect(url_for('register'))
        
        # Check if user exists
        if User.query.filter_by(email=email).first():
            flash('Email already registered')
            return redirect(url_for('register'))
        
        try:
            # Create new user
            hashed_password = generate_password_hash(password)
            user = User(name=name, email=email, password=hashed_password)
            db.session.add(user)
            db.session.commit()
            
            # Log in the new user
            login_user(user)
            flash('Registration successful!')
            return redirect(url_for('index'))
            
        except Exception as e:
            app.logger.error(f"Registration error: {str(e)}")
            db.session.rollback()
            flash('Registration failed. Please try again.')
            return redirect(url_for('register'))
    
    return render_template('register.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/add_sheet', methods=['POST'])
@login_required
def add_sheet():
    try:
        sheet_name = request.form.get('sheet_name', '').strip()
        data_raw = request.form.get('data', '').strip()
        
        if not sheet_name:
            flash('Sheet name is required!', 'error')
            return redirect(url_for('index'))
            
        if Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first():
            flash(f'Sheet "{sheet_name}" already exists!', 'error')
            return redirect(url_for('index'))

        # Parse data and ensure it's in the correct format
        data = []
        for row in data_raw.split(';'):
            row = row.strip()
            if row:
                row_data = []
                for cell in row.split(','):
                    cell = cell.strip()
                    # Preserve formulas
                    if cell.startswith('='):
                        row_data.append(cell)
                    else:
                        # Try to convert to number if possible
                        try:
                            if '.' in cell:
                                value = float(cell)
                            else:
                                value = int(cell)
                            row_data.append(value)
                        except ValueError:
                            row_data.append(cell)
                data.append(row_data)

        if not data:
            flash('No data provided!', 'error')
            return redirect(url_for('index'))

        # Store data as JSON string
        new_sheet = Sheet(
            sheet_name=sheet_name,
            data=json.dumps(data),
            user_id=current_user.id
        )
        db.session.add(new_sheet)
        db.session.commit()

        flash(f'Sheet "{sheet_name}" added!', 'success')
        return redirect(url_for('index'))

    except Exception as e:
        app.logger.error(f"Error adding sheet: {str(e)}")
        flash(f'Error adding sheet: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/remove_sheet/<sheet_name>', methods=['POST'])
@login_required
def remove_sheet(sheet_name):
    sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first()
    
    if sheet:
        db.session.delete(sheet)
        db.session.commit()
        flash(f'Sheet "{sheet_name}" removed successfully!', 'info')
    else:
        flash(f'No sheet named "{sheet_name}" found.', 'error')
    return redirect(url_for('index'))

@app.route('/reorder_sheets', methods=['POST'])
@login_required
def reorder_sheets():
    global sheet_order
    new_order = request.form.getlist('order[]')
    if new_order:
        sheet_order = new_order
        flash('Sheet order updated successfully!', 'success')
        return redirect(url_for('index'))
    flash('Failed to update sheet order.', 'error')
    return redirect(url_for('index'))

@app.route('/add_custom_template', methods=['POST'])
@login_required
def add_custom_template():
    try:
        template_id = request.form.get('template_id', '').strip()
        name = request.form.get('name', '').strip()
        default_sheet_name = request.form.get('default_sheet_name', '').strip()
        headers_raw = request.form.get('headers', '').strip()
        sample_data_raw = request.form.get('sample_data', '').strip()
        
        if not all([template_id, name, default_sheet_name, headers_raw]):
            flash("Missing required fields for custom template.", "error")
            return redirect(url_for('index'))
        
        formatting = {
            "header_bg_color": request.form.get('header_bg_color', "#FFFFFF"),
            "title_font_size": int(request.form.get('title_font_size', 12)),
            "title_bold": request.form.get('title_bold', 'off') == 'on',
            "cell_font_size": int(request.form.get('cell_font_size', 11)),
            "cell_alignment": request.form.get('cell_alignment', 'left'),
            "cell_border_color": request.form.get('cell_border_color', "#D3D3D3"),
            "text_wrap": request.form.get('text_wrap', 'off') == 'on',
            "number_format": request.form.get('number_format', 'General'),
            "merge_cells": request.form.get('merge_cells', ''),
            "freeze_panes": request.form.get('freeze_panes', '')
        }
        
        headers = [h.strip() for h in headers_raw.split(',') if h.strip()]
        sample_data = []
        if sample_data_raw:
            for row in sample_data_raw.split('\n'):
                if row.strip():
                    values = [v.strip() for v in row.split('\t')]
                    sample_data.append(values)
        
        new_template = {
            "id": template_id,
            "name": name,
            "default_sheet_name": default_sheet_name,
            "headers": headers,
            "sample_data": sample_data,
            "formatting": formatting
        }
        
        custom_templates.append(new_template)
        flash(f'Custom template "{name}" added successfully!', 'success')
        
    except Exception as e:
        app.logger.error(f"Error adding template: {str(e)}")
        flash('Error adding template. Please check your inputs.', 'error')
    
    return redirect(url_for('index'))

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    try:
        if 'file' not in request.files:
            app.logger.error("No file part in request")
            flash('No file submitted', 'error')
            return redirect(url_for('index'))
            
        file = request.files['file']
        
        if file.filename == '':
            app.logger.error("Empty filename submitted")
            flash('No file selected', 'error')
            return redirect(url_for('index'))

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            # Log file details
            app.logger.info(f"Processing file: {filename}")
            app.logger.info(f"File size: {os.fstat(file.fileno()).st_size} bytes")
            
            file.save(filepath)
            
            # Verify file was saved
            if not os.path.exists(filepath):
                app.logger.error(f"File not saved to {filepath}")
                flash('Error saving file', 'error')
                return redirect(url_for('index'))
                
            # Read and validate Excel data
            df = pd.read_excel(filepath, sheet_name=None)
            
            if not df:
                app.logger.error("No data found in Excel file")
                flash('Excel file appears to be empty', 'error')
                return redirect(url_for('index'))
                
            sheets_data = []
            sheet_order = []
            
            for sheet_name, sheet_df in df.items():
                if sheet_df.empty:
                    app.logger.warning(f"Sheet '{sheet_name}' is empty")
                    continue
                    
                sheet_data = {
                    'sheet_name': sheet_name,
                    'data': [sheet_df.columns.tolist()] + sheet_df.values.tolist()
                }
                sheets_data.append(sheet_data)
                sheet_order.append(sheet_name)
            
            if not sheets_data:
                app.logger.error("No valid data found in any sheet")
                flash('No valid data found in Excel file', 'error')
                return redirect(url_for('index'))
                
            # Store the data in session or database
            # Add code here to persist the data
                
            flash(f'Successfully processed {len(sheets_data)} sheets', 'success')
            return redirect(url_for('index'))
            
    except Exception as e:
        app.logger.exception("Error processing upload")
        flash(f'Error processing file: {str(e)}', 'error')
        
    return redirect(url_for('index'))

@app.route('/download_excel')
@login_required
def download_excel():
    try:
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        data_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        number_format = workbook.add_format({
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'  # Format for numbers
        })
        
        # Get user's sheets from database
        user_sheets = Sheet.query.filter_by(user_id=current_user.id).all()
        
        if not user_sheets:
            flash('No data available to download', 'error')
            return redirect(url_for('index'))

        # Process each sheet
        for sheet in user_sheets:
            worksheet = workbook.add_worksheet(sheet.sheet_name[:31])
            
            try:
                # Parse JSON data
                sheet_data = json.loads(sheet.data) if isinstance(sheet.data, str) else sheet.data
                
                if not sheet_data or not isinstance(sheet_data, list):
                    continue
                
                # Get headers and identify numeric columns
                headers = sheet_data[0] if sheet_data else []
                numeric_cols = []
                formula_cols = []
                
                # Write headers and analyze data types
                for col_idx, header in enumerate(headers):
                    worksheet.write(0, col_idx, header, header_format)
                    
                    # Check if column contains numbers or formulas
                    if len(sheet_data) > 1:
                        sample_value = sheet_data[1][col_idx] if col_idx < len(sheet_data[1]) else None
                        if isinstance(sample_value, (int, float)) or (isinstance(sample_value, str) and sample_value.replace('.', '').isdigit()):
                            numeric_cols.append(col_idx)
                        elif isinstance(sample_value, str) and sample_value.startswith('='):
                            formula_cols.append(col_idx)
                
                # Write data rows
                for row_idx, row_data in enumerate(sheet_data[1:], start=1):
                    for col_idx, cell_value in enumerate(row_data):
                        # Handle different cell types
                        if col_idx in numeric_cols:
                            try:
                                # Convert to float if possible
                                value = float(cell_value) if cell_value else 0
                                worksheet.write_number(row_idx, col_idx, value, number_format)
                            except (ValueError, TypeError):
                                worksheet.write(row_idx, col_idx, cell_value, data_format)
                        elif col_idx in formula_cols:
                            if isinstance(cell_value, str) and cell_value.startswith('='):
                                worksheet.write_formula(row_idx, col_idx, cell_value, number_format)
                            else:
                                worksheet.write(row_idx, col_idx, cell_value, data_format)
                        else:
                            worksheet.write(row_idx, col_idx, cell_value, data_format)
                
                # Auto-adjust column widths
                for col_idx, _ in enumerate(headers):
                    max_length = len(str(headers[col_idx]))
                    for row_idx in range(1, len(sheet_data)):
                        cell_value = str(sheet_data[row_idx][col_idx])
                        max_length = max(max_length, len(cell_value))
                    worksheet.set_column(col_idx, col_idx, max_length + 2)
                        
            except Exception as sheet_error:
                app.logger.error(f"Error processing sheet {sheet.sheet_name}: {str(sheet_error)}")
                continue
        
        workbook.close()
        output.seek(0)
        
        # Save download history
        history = DownloadHistory(
            user_id=current_user.id,
            filename=f'workbook_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
            sheet_data=json.dumps([{
                'sheet_name': sheet.sheet_name,
                'data': sheet.data
            } for sheet in user_sheets])
        )
        db.session.add(history)
        db.session.commit()

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'workbook_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

    except Exception as e:
        app.logger.error(f"Download error: {str(e)}")
        flash(f'Error generating Excel file: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/update_sheet_format/<sheet_name>', methods=['POST'])
@login_required
def update_sheet_format(sheet_name):
    global sheets_data
    sheet = next((s for s in sheets_data if s['sheet_name'] == sheet_name), None)
    if not sheet:
        flash(f'Sheet "{sheet_name}" not found.', 'error')
        return redirect(url_for('index'))
    
    try:
        # Update formatting
        sheet['formatting'].update({
            'font': {
                'name': request.form.get('font_name', 'Arial'),
                'size': int(request.form.get('font_size', 11)),
                'bold': request.form.get('font_bold', 'off') == 'on',
                'italic': request.form.get('font_italic', 'off') == 'on',
                'color': request.form.get('font_color', '#000000')
            },
            'cell': {
                'alignment': request.form.get('cell_alignment', 'left'),
                'vertical_alignment': request.form.get('vertical_alignment', 'center'),
                'wrap_text': request.form.get('text_wrap', 'off') == 'on',
                'bg_color': request.form.get('cell_bg_color', '#FFFFFF')
            },
            'header': {
                'bg_color': request.form.get('header_bg_color', '#F0E68C'),
                'bold': request.form.get('header_bold', 'on') == 'on',
                'border': request.form.get('header_border', 'on') == 'on'
            },
            'border': {
                'style': request.form.get('border_style', 'thin'),
                'color': request.form.get('border_color', '#000000'),
                'all_borders': request.form.get('all_borders', 'off') == 'on',
                'outline': request.form.get('outline_only', 'off') == 'on'
            },
            'column_width': float(request.form.get('column_width', 15)),
            'row_height': float(request.form.get('row_height', 15))
        })
        
        flash(f'Formatting updated for sheet "{sheet_name}".', 'success')
    except Exception as e:
        app.logger.error(f"Error updating format for sheet {sheet_name}: {str(e)}")
        flash('Error updating sheet formatting.', 'error')
    return redirect(url_for('index'))

@app.route('/update_sheet_data', methods=['POST'])
@login_required
def update_sheet_data():
    global sheets_data
    data = request.get_json()
    sheet_name = data.get('sheet_name')
    updated_data = data.get('data')

    sheet = next((s for s in sheets_data if s['sheet_name'] == sheet_name), None)
    if not sheet:
        return jsonify({'success': False, 'error': f'Sheet "{sheet_name}" not found.'})
    try:
        sheet['data'] = [[cell.strip() for cell in row] for row in updated_data]
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/history')
@login_required
def view_history():
    downloads = DownloadHistory.query.filter_by(user_id=current_user.id)\
               .order_by(DownloadHistory.downloaded_at.desc()).all()
    return render_template('history.html', downloads=downloads)

@app.route('/add_to_templates/<int:download_id>', methods=['POST'])
@login_required
def add_to_templates(download_id):
    download = DownloadHistory.query.get_or_404(download_id)
    if download.user_id != current_user.id:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 403
    
    try:
        # Parse sheet data from history
        sheet_data_json = json.loads(download.sheet_data) if isinstance(download.sheet_data, str) else download.sheet_data
        
        # Add each sheet to custom templates
        for sheet_item in sheet_data_json:
            sheet_data = json.loads(sheet_item['data']) if isinstance(sheet_item['data'], str) else sheet_item['data']
            
            # Create template from sheet
            template_id = f"template_{download.id}_{len(custom_templates)}"
            template_name = f"Template from {download.filename}"
            
            # Get headers and sample data
            headers = sheet_data[0] if sheet_data and len(sheet_data) > 0 else []
            sample_data = sheet_data[1:3] if sheet_data and len(sheet_data) > 1 else []
            
            # Add to templates
            custom_templates.append({
                'id': template_id,
                'name': template_name,
                'default_sheet_name': sheet_item['sheet_name'],
                'headers': headers,
                'sample_data': sample_data,
                'formatting': {
                    'header_bg_color': '#F0E68C',  # Default yellow
                    'title_font_size': 14,
                    'title_bold': True
                }
            })
        
        return jsonify({'success': True})
    except Exception as e:
        app.logger.error(f"Error adding template: {str(e)}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/profile')
@login_required
def profile():
    return render_template('profile.html', user=current_user)

if __name__ == '__main__':
    # init_db()  # Only use this for initial setup or debugging!
    app.run(debug=True)