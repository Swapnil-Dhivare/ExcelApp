from flask import request, jsonify, redirect, url_for, render_template, flash
from flask_login import login_required, current_user
from flask_wtf.csrf import generate_csrf
from datetime import datetime
from . import app
from .models import db, Sheet
import json

# ...existing code...

@app.route('/add_sheet', methods=['GET', 'POST'])
@login_required
def add_sheet():
    # Get built-in templates and custom templates (if any)
    built_in_templates = []  # You can define your built-in templates here
    custom_templates = []    # You can fetch user's custom templates if implemented
    
    if request.method == 'POST':
        try:
            # Get form data
            sheet_name = request.form.get('sheet_name')
            if not sheet_name:
                flash('Sheet name is required', 'error')
                return render_template('add_sheet.html', 
                                      built_in_templates=built_in_templates, 
                                      custom_templates=custom_templates)
            
            # Check if sheet name already exists for this user
            existing_sheet = Sheet.query.filter_by(user_id=current_user.id, sheet_name=sheet_name).first()
            if existing_sheet:
                flash('A sheet with this name already exists', 'error')
                return render_template('add_sheet.html', 
                                      built_in_templates=built_in_templates, 
                                      custom_templates=custom_templates)
            
            # Create empty data structure for new sheet
            # Default structure with one row and one column
            data = [["Header"], ["Cell"]]
            
            # Check if template was selected
            template_id = request.form.get('template_id')
            if template_id:
                # Logic to load template data would go here
                pass
                
            # Check if headers were provided
            headers = request.form.get('headers')
            if headers:
                header_list = [h.strip() for h in headers.split(',')]
                if header_list:
                    data[0] = header_list
                    # Add empty row with same number of columns
                    data[1] = ["" for _ in range(len(header_list))]
            
            # Check if data was provided
            input_data = request.form.get('data')
            if input_data:
                try:
                    rows = input_data.strip().split(';')
                    sheet_data = []
                    
                    # Keep the headers
                    sheet_data.append(data[0])
                    
                    # Process each row of data
                    for row in rows:
                        if row.strip():
                            cells = row.split(',')
                            # Pad or truncate to match header length
                            while len(cells) < len(data[0]):
                                cells.append("")
                            sheet_data.append(cells[:len(data[0])])
                    
                    if len(sheet_data) > 1:  # If we have data rows
                        data = sheet_data
                except Exception as e:
                    flash(f'Error parsing data: {str(e)}', 'error')
            
            # Create new sheet in database
            new_sheet = Sheet(
                sheet_name=sheet_name,
                data=data,
                user_id=current_user.id,
                created_at=datetime.utcnow(),
                updated_at=datetime.utcnow()
            )
            
            db.session.add(new_sheet)
            db.session.commit()
            
            flash('Sheet created successfully!', 'success')
            return redirect(url_for('index'))
        except Exception as e:
            db.session.rollback()
            flash(f'Error: {str(e)}', 'error')
    
    # For GET request, render the form
    return render_template('add_sheet.html', 
                          built_in_templates=built_in_templates, 
                          custom_templates=custom_templates)

# Add this route to provide CSRF token via API if needed
@app.route('/get_csrf_token', methods=['GET'])
def get_csrf_token():
    return jsonify({'csrf_token': generate_csrf()})

# ...existing code...