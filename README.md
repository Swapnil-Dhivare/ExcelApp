📊 Excel Generator Web App
A Flask-based web application that lets users create, customize, and download Excel files with advanced formatting.

✨ Features
✅ Dynamic Sheet Creation

Add/remove sheets with custom names

Input data via forms or CSV-style text

✅ Professional Excel Formatting

Custom fonts, colors, borders, and alignment

Cell merging, freeze panes, text wrapping

Data validation (dropdown lists)

Conditional formatting (data bars, color scales)

✅ Pre-Built & Custom Templates

Built-in templates (Front Page, Summary Sheet, Data Sheet)

Save your own templates for reuse

✅ Import/Export

Upload existing Excel files (auto-converts to app format)

Download multi-sheet workbooks as .xlsx

✅ User-Friendly UI

Drag-and-drop sheet reordering

Real-time previews

Responsive design (works on desktop/tablet)

🛠️ Tech Stack
Backend: Python + Flask

Excel Engine: Pandas + XlsxWriter

Frontend: Bootstrap 5 + Jinja2

Deployment: Ready for Heroku/Docker (add your own config)

🚀 Use Cases
Business reports

Data dashboards

Academic templates

Personal finance trackers

📥 Installation
bash
Copy
git clone https://github.com/yourusername/excel-generator.git
cd excel-generator
pip install -r requirements.txt
flask run
🔧 How to Extend
Add charts/graphs with openpyxl

Integrate user authentication

Connect to databases (SQLite, PostgreSQL)
