# Student Certificate/Receipt Generator Web Application

## Overview
A Python Flask web application to automate the generation of student certificates or receipts by extracting data from Excel files and overlaying it onto template images.

## Features
- Upload Excel files (.xlsx, .xls) with student data
- Upload background template images (PNG, JPG, JPEG)
- Configure text and image overlay positions
- Batch generate certificates/receipts as images
- Download results as a ZIP archive

## Tech Stack
- Python 3.8+
- Flask
- pandas, openpyxl
- Pillow (PIL)
- Bootstrap (frontend)

## Setup
1. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```
2. Run the app:
   ```sh
   export FLASK_APP=app
   flask run
   ```
3. Open your browser at http://127.0.0.1:5001/

## Folder Structure
- `app/` - Flask application code
- `templates/` - HTML templates
- `static/` - Static files (CSS, JS, images)
- `output/` - Generated certificates/receipts
- `config/` - Configuration files

## To Do
- Implement Excel parsing and image overlay logic
- Add configuration UI for positioning
- Add batch processing and ZIP export

---
