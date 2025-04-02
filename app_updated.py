from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
import pandas as pd
import io
import logging
import xlsxwriter
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'
app.logger.setLevel(logging.DEBUG)

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

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
    }
]

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    sorted_sheets = []
    for sheet_name in sheet_order:
        sheet = next((s for s in sheets_data if s['sheet_name'] == sheet_name), None)
        if sheet:
            sorted_sheets.append(sheet)
    
    for sheet in sheets_data:
        if sheet['sheet_name'] not in sheet_order:
            sorted_sheets.append(sheet)
            sheet_order.append(sheet['sheet_name'])
    
    return render_template(
        'index.html',
        sheets=sorted_sheets,
        built_in_templates=built_in_templates,
        custom_templates=custom_templates
    )

@app.route('/add_sheet', methods=['POST'])
def add_sheet():
    try:
        sheet_name = request.form.get('sheet_name', '').strip()
        template_id = request.form.get('template_id', '')
        
        if any(sheet['sheet_name'] == sheet_name for sheet in sheets_data):
            flash(f'Sheet "{sheet_name}" already exists!', 'error')
            return redirect(url_for('index'))
        
        if template_id:
            template = next((t for t in built_in_templates + custom_templates if t['id'] == template_id), None)
            if template:
                new_sheet = {
                    "sheet_name": sheet_name or template['default_sheet_name'],
                    "headers": template['headers'].copy(),
                    "data": [row.copy() for row in template['sample_data']],
                    "formatting": template['formatting'].copy()
                }
                sheets_data.append(new_sheet)
                sheet_order.append(new_sheet['sheet_name'])
                flash(f'Sheet "{new_sheet["sheet_name"]}" created from template!', 'success')
                return redirect(url_for('index'))
        
        # Process manual sheet creation
        headers = request.form.get('headers', '')
        data_raw = request.form.get('data', '')
        
        formatting = {
            "title_font_size": int(request.form.get('title_font_size', 12)),
            "title_bold": request.form.get('title_bold', 'off') == 'on',
            "header_bg_color": request.form.get('header_bg_color', "#F0E68C"),
            "cell_font_size": int(request.form.get('cell_font_size', 11)),
            "cell_alignment": request.form.get('cell_alignment', 'left'),
            "cell_border_color": request.form.get('cell_border_color', "#D3D3D3"),
            "text_wrap": request.form.get('text_wrap', 'off') == 'on',
            "number_format": request.form.get('number_format', 'General'),
            "merge_cells": request.form.get('merge_cells', ''),
            "freeze_panes": request.form.get('freeze_panes', '')
        }
        
        # Data validation
        validation = {}
        val_types = request.form.getlist('val_type[]')
        val_columns = request.form.getlist('val_column[]')
        val_options = request.form.getlist('val_options[]')
        
        for i, col in enumerate(val_columns):
            if col and val_types[i] and val_types[i] != 'none':
                validation[col] = {
                    "type": val_types[i],
                    "options": [opt.strip() for opt in val_options[i].split(',')] if val_options[i] else []
                }
        
        if validation:
            formatting['validation'] = validation
        
        headers_list = [h.strip() for h in headers.split(',') if h.strip()] if headers else []
        rows = []
        if data_raw:
            for row in data_raw.split('\n'):
                if row.strip():
                    values = [v.strip() for v in row.split('\t')]
                    if len(values) < len(headers_list):
                        values += [''] * (len(headers_list) - len(values))
                    rows.append(values)
        
        if sheet_name:
            new_sheet = {
                "sheet_name": sheet_name,
                "headers": headers_list,
                "data": rows,
                "formatting": formatting
            }
            sheets_data.append(new_sheet)
            sheet_order.append(sheet_name)
            flash(f'Sheet "{sheet_name}" added successfully!', 'success')
        
    except Exception as e:
        app.logger.error(f"Error adding sheet: {str(e)}")
        flash('Error adding sheet. Please check your inputs.', 'error')
    
    return redirect(url_for('index'))

@app.route('/remove_sheet/<sheet_name>', methods=['POST'])
def remove_sheet(sheet_name):
    global sheets_data, sheet_order
    original_count = len(sheets_data)
    sheets_data = [sheet for sheet in sheets_data if sheet["sheet_name"] != sheet_name]
    sheet_order = [name for name in sheet_order if name != sheet_name]
    
    if len(sheets_data) < original_count:
        flash(f'Sheet "{sheet_name}" removed successfully!', 'info')
    else:
        flash(f'No sheet named "{sheet_name}" found.', 'error')
    return redirect(url_for('index'))

@app.route('/reorder_sheets', methods=['POST'])
def reorder_sheets():
    global sheet_order
    new_order = request.json.get('order', [])
    if new_order:
        sheet_order = new_order
        return jsonify(success=True)
    return jsonify(success=False), 400

@app.route('/add_custom_template', methods=['POST'])
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
def upload_file():
    if 'file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            xls = pd.ExcelFile(filepath)
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                sheets_data.append({
                    "sheet_name": sheet_name,
                    "headers": df.columns.tolist(),
                    "data": df.values.tolist(),
                    "formatting": {
                        "header_bg_color": "#F0E68C",
                        "cell_alignment": "left"
                    }
                })
                sheet_order.append(sheet_name)
            
            flash(f'Successfully imported {len(xls.sheet_names)} sheets from {filename}', 'success')
            
        except Exception as e:
            app.logger.error(f"Error processing file: {str(e)}")
            flash('Error processing Excel file', 'error')
        
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
    else:
        flash('Allowed file types are .xlsx, .xls', 'error')
    
    return redirect(url_for('index'))

@app.route('/download_excel')
def download_excel():
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            for sheet_name in sheet_order:
                sheet = next((s for s in sheets_data if s['sheet_name'] == sheet_name), None)
                if not sheet:
                    continue
                
                df = pd.DataFrame(sheet['data'], columns=sheet['headers']) if sheet['headers'] else pd.DataFrame(sheet['data'])
                df.to_excel(writer, sheet_name=sheet['sheet_name'], index=False)
                worksheet = writer.sheets[sheet['sheet_name']]
                fmt = sheet.get("formatting", {})
                
                # Apply formatting
                header_format = workbook.add_format({
                    'bold': fmt.get("title_bold", False),
                    'font_size': fmt.get("title_font_size", 12),
                    'bg_color': fmt.get("header_bg_color", "#F0E68C"),
                    'align': fmt.get("cell_alignment", "left"),
                    'valign': 'vcenter',
                    'border': 1,
                    'border_color': fmt.get("cell_border_color", "#D3D3D3"),
                    'text_wrap': fmt.get("text_wrap", False)
                })
                
                cell_format = workbook.add_format({
                    'font_size': fmt.get("cell_font_size", 11),
                    'align': fmt.get("cell_alignment", "left"),
                    'valign': 'vcenter',
                    'border': 1,
                    'border_color': fmt.get("cell_border_color", "#D3D3D3"),
                    'text_wrap': fmt.get("text_wrap", False),
                    'num_format': fmt.get("number_format", "General")
                })
                
                # Write data
                for col_num, header in enumerate(sheet['headers']):
                    worksheet.write(0, col_num, header, header_format)
                    worksheet.set_column(col_num, col_num, 20)
                
                for r, row in enumerate(sheet['data']):
                    for c, val in enumerate(row[:len(sheet['headers'])] if sheet['headers'] else row):
                        if isinstance(val, str) and val.startswith('='):
                            worksheet.write_formula(r + 1, c, val, cell_format)
                        else:
                            worksheet.write(r + 1, c, val, cell_format)
                
                # Advanced formatting
                if fmt.get("merge_cells"):
                    for merge_range in fmt["merge_cells"].split(';'):
                        if merge_range.strip():
                            worksheet.merge_range(merge_range.strip(), "", cell_format)
                
                if fmt.get("freeze_panes"):
                    worksheet.freeze_panes(fmt["freeze_panes"])
                
                if fmt.get("validation"):
                    for col_name, val in fmt["validation"].items():
                        if col_name in sheet['headers'] and val['type'] == 'list':
                            col_idx = sheet['headers'].index(col_name)
                            worksheet.data_validation(
                                1, col_idx, len(sheet['data']), col_idx,
                                {
                                    'validate': 'list',
                                    'source': val['options'],
                                    'input_title': 'Select:',
                                    'input_message': 'Choose from dropdown'
                                }
                            )
                
                if fmt.get("conditional_formatting"):
                    for range_, rules in fmt["conditional_formatting"].items():
                        if rules['type'] == 'data_bar':
                            worksheet.conditional_format(range_, {
                                'type': 'data_bar',
                                'bar_color': rules.get('bar_color', '#63C384'),
                                'min_type': rules.get('min_type', 'num'),
                                'min_value': rules.get('min_value', 0),
                                'max_type': rules.get('max_type', 'num'),
                                'max_value': rules.get('max_value', 100)
                            })
        
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name="workbook.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except Exception as e:
        app.logger.error(f"Error generating Excel: {str(e)}")
        flash('Error generating Excel file', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)