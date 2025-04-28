import os
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from templates import get_built_in_templates
from flask_wtf.csrf import CSRFProtect  # Add this import for CSRF protection

# Initialize Flask app
app = Flask(__name__)

# Load configuration from config.py
app.config.from_pyfile('config.py')

# Initialize CSRF protection
csrf = CSRFProtect(app)

# Initialize database
db = SQLAlchemy(app)

# Initialize LoginManager
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# User model
class User(UserMixin, db.Model):
    __tablename__ = 'user'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255))  # Increase size from 128 to 255
    name = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relationship with Sheet model
    sheets = db.relationship('Sheet', backref='user', lazy=True, cascade="all, delete-orphan")
    
    # Relationship with DownloadHistory model
    downloads = db.relationship('DownloadHistory', backref='user', lazy=True)
    
    def set_password(self, password):
        # Use generate_password_hash with specified method and salt length
        self.password_hash = generate_password_hash(
            password, 
            method='pbkdf2:sha256', 
            salt_length=8
        )
        
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
        
    def __repr__(self):
        return f'<User {self.username}>'

# Sheet model
class Sheet(db.Model):
    __tablename__ = 'sheet'
    
    id = db.Column(db.Integer, primary_key=True)
    sheet_name = db.Column(db.String(100), nullable=False)
    data = db.Column(db.JSON, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationship with DownloadHistory model
    downloads = db.relationship('DownloadHistory', backref='sheet', lazy=True)
    
    def __repr__(self):
        return f'<Sheet {self.sheet_name}>'

# Download History model
class DownloadHistory(db.Model):
    __tablename__ = 'download_history'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    sheet_id = db.Column(db.Integer, db.ForeignKey('sheet.id'), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    format = db.Column(db.String(20), nullable=False)  # xlsx, csv, pdf
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    file_path = db.Column(db.String(255), nullable=True)  # Path to saved file if stored

# User loader function for Flask-Login
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# Routes for your application
@app.route('/')
def index():
    if current_user.is_authenticated:
        return render_template('index.html')
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for('index'))
        
        flash('Invalid username or password')
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        name = request.form.get('name')
        
        # Check if username or email already exists
        if User.query.filter_by(username=username).first():
            flash('Username already exists')
            return redirect(url_for('register'))
        
        if User.query.filter_by(email=email).first():
            flash('Email already exists')
            return redirect(url_for('register'))
        
        # Create new user
        user = User(username=username, email=email, name=name)
        user.set_password(password)
        
        db.session.add(user)
        db.session.commit()
        
        flash('Registration successful! Please login.')
        return redirect(url_for('login'))
    
    return render_template('register.html')

@app.route('/add_sheet', methods=['GET', 'POST'])
@login_required
def add_sheet():
    if request.method == 'POST':
        sheet_name = request.form.get('sheet_name')
        data = request.form.get('data')
        
        if not sheet_name:
            flash('Sheet name is required.', 'error')
            return redirect(url_for('add_sheet'))
            
        # Check if sheet name already exists for this user
        existing_sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first()
        
        if existing_sheet:
            flash('A sheet with this name already exists.', 'error')
            return redirect(url_for('add_sheet'))
            
        # Parse data if provided
        sheet_data = []
        if data:
            rows = data.strip().split(';')
            for row in rows:
                if row:
                    sheet_data.append(row.split(','))
        else:
            # Default empty sheet with one header row and 5 data rows
            sheet_data = [[''] * 5 for _ in range(6)]
            
        # Create new sheet
        new_sheet = Sheet(
            sheet_name=sheet_name,
            data=sheet_data,
            user_id=current_user.id
        )
        
        db.session.add(new_sheet)
        db.session.commit()
        
        flash(f'Sheet "{sheet_name}" created successfully!', 'success')
        return redirect(url_for('edit_sheet', sheet_name=sheet_name))
    
    # GET request - render the form with templates
    built_in_templates = get_built_in_templates()
    
    # Add a sample_data_string property to each template
    for template in built_in_templates:
        if hasattr(template, 'sample_data') and template.sample_data:
            template.sample_data_string = template.sample_data_string
        else:
            template.sample_data_string = ""
    
    # Add custom templates if needed
    custom_templates = []  # You can add custom template logic here
    
    return render_template('add_sheet.html', 
                          built_in_templates=built_in_templates,
                          custom_templates=custom_templates)

@app.route('/edit_sheet/<sheet_name>')
@login_required
def edit_sheet(sheet_name):
    sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first()
    
    if not sheet:
        flash('Sheet not found.', 'error')
        return redirect(url_for('index'))
    
    return render_template('sheet_editor.html', sheet=sheet, data=sheet.data)

@app.route('/update_sheet/<sheet_name>', methods=['POST'])
@login_required
def update_sheet(sheet_name):
    sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first()
    
    if not sheet:
        return jsonify({'success': False, 'message': 'Sheet not found'})
    
    data = request.json.get('data')
    if not data:
        return jsonify({'success': False, 'message': 'No data provided'})
    
    sheet.data = data
    sheet.updated_at = datetime.utcnow()
    db.session.commit()
    
    return jsonify({'success': True, 'message': 'Sheet updated successfully'})

@app.route('/delete_sheet/<sheet_name>', methods=['POST'])
@login_required
def delete_sheet(sheet_name):
    sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first()
    
    if not sheet:
        return jsonify({'success': False, 'message': 'Sheet not found'})
    
    db.session.delete(sheet)
    db.session.commit()
    
    return jsonify({'success': True, 'message': 'Sheet deleted successfully'})

@app.route('/get_template/<template_id>')
@login_required
def get_template(template_id):
    templates = get_built_in_templates()
    template = next((t for t in templates if str(t.id) == template_id), None)
    
    if not template:
        return jsonify({'success': False, 'error': 'Template not found'})
    
    # Convert headers to string
    headers = ','.join(template.headers) if template.headers else ''
    
    # Use the pre-computed sample_data_string
    sample_data = template.sample_data_string
    
    return jsonify({
        'success': True,
        'headers': headers,
        'sample_data': sample_data
    })

@app.route('/history')
@login_required
def view_history():
    """View the history of sheets created by the current user."""
    # Get all sheets for the current user, ordered by creation date
    sheets = Sheet.query.filter_by(user_id=current_user.id).order_by(Sheet.created_at.desc()).all()
    
    return render_template('history.html', sheets=sheets)

@app.route('/download_history')
@login_required
def download_history():
    """View download history for the current user"""
    downloads = DownloadHistory.query.filter_by(user_id=current_user.id)\
        .order_by(DownloadHistory.created_at.desc()).all()
    return render_template('download_history.html', downloads=downloads)

@app.route('/manage_templates')
@login_required
def manage_templates():
    """Manage templates including adding from download history"""
    # Get built-in templates
    built_in_templates = get_built_in_templates()
    
    # Get user's download history for potential templates
    downloads = DownloadHistory.query.filter_by(user_id=current_user.id)\
        .order_by(DownloadHistory.created_at.desc()).all()
    
    return render_template('manage_templates.html', 
                          built_in_templates=built_in_templates,
                          downloads=downloads)

@app.route('/delete_template/<template_id>', methods=['POST'])
@login_required
def delete_template(template_id):
    """Delete a custom template"""
    # Get custom templates from session
    custom_templates = session.get('custom_templates', [])
    
    # Find the template with matching ID
    found = False
    for i, template in enumerate(custom_templates):
        if template.get('id') == template_id:
            # Remove the template
            custom_templates.pop(i)
            found = True
            break
    
    if found:
        # Update the session
        session['custom_templates'] = custom_templates
        session.modified = True
        return jsonify({'success': True, 'message': 'Template deleted successfully'})
    else:
        return jsonify({'success': False, 'message': 'Template not found'})

@app.route('/user_profile')
@login_required
def user_profile():
    """View user profile"""
    # Count sheets and downloads
    sheet_count = Sheet.query.filter_by(user_id=current_user.id).count()
    download_count = DownloadHistory.query.filter_by(user_id=current_user.id).count()
    custom_template_count = len(session.get('custom_templates', []))
    
    return render_template('user_profile.html', 
                          sheet_count=sheet_count, 
                          download_count=download_count,
                          custom_template_count=custom_template_count)

@app.route('/user_settings')
@login_required
def user_settings():
    """User settings page"""
    return render_template('user_settings.html')

def track_download(sheet_id, filename, format_type, file_path=None):
    """Track a file download in history"""
    if current_user.is_authenticated:
        download = DownloadHistory(
            user_id=current_user.id,
            sheet_id=sheet_id,
            filename=filename,
            format=format_type,
            file_path=file_path
        )
        db.session.add(download)
        db.session.commit()
        return download
    return None

@app.route('/export/<sheet_name>/<format_type>')
@login_required
def export_sheet(sheet_name, format_type):
    """Export a sheet to various formats and track the download"""
    sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first_or_404()
    
    if format_type not in ['xlsx', 'csv', 'pdf']:
        flash('Unsupported format', 'error')
        return redirect(url_for('edit_sheet', sheet_name=sheet_name))
    
    # Generate the file (implementation depends on your export functions)
    filename = f"{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{format_type}"
    
    # Track the download
    download = track_download(sheet.id, filename, format_type)
    
    # Return the file download response
    # This is a placeholder - implement your actual file generation and serving logic
    if format_type == 'xlsx':
        # Generate Excel file
        pass
    elif format_type == 'csv':
        # Generate CSV file
        pass
    elif format_type == 'pdf':
        # Generate PDF file
        pass
    
    flash(f'File exported successfully as {format_type.upper()}', 'success')
    return redirect(url_for('edit_sheet', sheet_name=sheet_name))

@app.route('/add_to_templates/<int:download_id>')
@login_required
def add_to_templates(download_id):
    """Add a downloaded file as a template"""
    download = DownloadHistory.query.filter_by(id=download_id, user_id=current_user.id).first_or_404()
    
    # Get the sheet associated with this download
    sheet = download.sheet
    
    if not sheet:
        flash('Sheet not found', 'error')
        return redirect(url_for('manage_templates'))
    
    # Get the next available template ID
    template_id = str(int(datetime.now().timestamp()))
    
    # Create a new template from the sheet data
    new_template = {
        'id': template_id,
        'name': f"{sheet.sheet_name} Template",
        'headers': sheet.data[0] if sheet.data and len(sheet.data) > 0 else [],
        'sample_data': sheet.data[1:6] if sheet.data and len(sheet.data) > 1 else [],
        'default_sheet_name': sheet.sheet_name,
        'description': f"Template created from {download.filename}",
        'user_id': current_user.id,
        'is_custom': True
    }
    
    # Save the custom template (implementation depends on how you store custom templates)
    save_custom_template(new_template)
    
    flash('Template created successfully', 'success')
    return redirect(url_for('manage_templates'))

def save_custom_template(template_data):
    """Save a custom template to the database or file system"""
    # This is a placeholder function
    # Implement based on how you want to store custom templates
    # You might want to create a CustomTemplate model and save to the database,
    # or store templates in a JSON file, etc.
    
    # For now, we'll store in session as an example
    if 'custom_templates' not in session:
        session['custom_templates'] = []
    
    session['custom_templates'].append(template_data)
    session.modified = True

# Custom filter for column letters
@app.template_filter('column_letter')
def column_letter(num):
    """Convert a number to Excel-style column letter."""
    result = ""
    while num >= 0:
        result = chr(65 + num % 26) + result
        num = num // 26 - 1
        if num < 0:
            break
    return result

# Create database tables if they don't exist
with app.app_context():
    db.create_all()
    
    # Create a demo user if none exists
    if not User.query.filter_by(username='demo').first():
        demo_user = User(username='demo', email='demo@example.com', name='Demo User')
        demo_user.set_password('password')
        db.session.add(demo_user)
        db.session.commit()

if __name__ == '__main__':
    app.run(debug=True)