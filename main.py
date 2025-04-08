from flask import Flask, render_template, request, send_file
import os
import re
import pdfplumber
import arabic_reshaper
import xlsxwriter
from bidi.algorithm import get_display

app = Flask(__name__)

uploaded_file_paths = []

app.config['UPLOAD_FOLDER'] = 'process'  # Set the upload folder

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    global uploaded_file_paths
    uploaded_file_paths = []
    uploaded_files = request.files.getlist('files[]')
    for uploaded_file in uploaded_files:
        if uploaded_file.filename != '':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
            uploaded_file.save(file_path)
            uploaded_file_paths.append(file_path)
    process_files()
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], 'sheet.xlsx'), as_attachment=True), delete()

def extract_visa_info_from_text(text, file_name):
    # Extract the name directly from the file name
    name = os.path.splitext(os.path.basename(file_name))[0]
    
    country_match = re.search(r'Nationality\s+([A-Za-z ]+)', text)
    country = country_match.group(1).strip() if country_match else None
    
    # Adjusting the regex for Passport Number to allow alphanumeric values
    passport_match = re.search(r'(?:Passport No.|PassportNo.|رقم الجواز)\s*([A-Z0-9]+)', text)
    passport_number = passport_match.group(1).strip() if passport_match else None
    
    valid_from_match = re.search(r'Valid From\s+(\d{2}/\d{2}/\d{4})', text)
    valid_from = valid_from_match.group(1).strip() if valid_from_match else None
    
    valid_until_match = re.search(r'Valid Until\s+(\d{2}/\d{2}/\d{4})', text)
    valid_until = valid_until_match.group(1).strip() if valid_until_match else None
    
    # Improved extraction for 'Duration of Stay'
    duration_of_stay = None
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "Duration of Stay" in line:
            # Try to find a numeric value in the same line or the next one
            duration_line = line + (lines[i+1] if i+1 < len(lines) else "")
            duration_match = re.search(r'(\d+|[٠-٩]+)', duration_line)
            if duration_match:
                duration_of_stay = duration_match.group(1).strip()
                # Convert Arabic numerals to English if necessary
                arabic_to_english = {'٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4', 
                                     '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'}
                duration_of_stay = ''.join(arabic_to_english.get(c, c) for c in duration_of_stay)
            break
    
    entry_type_match = re.search(r'Entry Type\s+(Single|Multiple)', text)
    entry_type = entry_type_match.group(1).strip() if entry_type_match else None
    
    visa_no_match = re.search(r'Visa No.\s+(\d+)', text)
    visa_no = visa_no_match.group(1).strip() if visa_no_match else None
    
    return {
        "Name": name,  # Extracted from file name
        "Country": country,
        "Passport Number": passport_number,
        "Visa No": visa_no,
        "Valid From": valid_from,
        "Valid Until": valid_until,
        "Duration of Stay": duration_of_stay,
        "Entry Type": entry_type
    }

def process_files():
    workbook = xlsxwriter.Workbook(os.path.join(app.config['UPLOAD_FOLDER'], 'sheet.xlsx'))
    worksheet = workbook.add_worksheet()
    
    headers = ["Name", "Country", "Passport Number", "Visa No", "Valid From", "Valid Until", "Duration of Stay", "Entry Type"]
    for col_num, header in enumerate(headers, start=1):
        worksheet.write(0, col_num - 1, header)
    
    for row_num, file_path in enumerate(uploaded_file_paths, start=1):
        with pdfplumber.open(file_path) as pdf:
            first_page = pdf.pages[0]
            extracted_text = first_page.extract_text()
        
        reshaped_text = arabic_reshaper.reshape(extracted_text)
        bidi_text = get_display(reshaped_text)
        
        visa_info = extract_visa_info_from_text(bidi_text, file_path)
        
        worksheet.write(row_num, 0, visa_info["Name"])
        worksheet.write(row_num, 1, visa_info["Country"])
        worksheet.write(row_num, 2, visa_info["Passport Number"])
        worksheet.write(row_num, 3, visa_info["Visa No"])
        worksheet.write(row_num, 4, visa_info["Valid From"])
        worksheet.write(row_num, 5, visa_info["Valid Until"])
        worksheet.write(row_num, 6, visa_info["Duration of Stay"])
        worksheet.write(row_num, 7, visa_info["Entry Type"])
    
    workbook.close()

def delete():
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)  # Create 'process' folder if it doesn't exist
    app.run(debug=True, host='0.0.0.0', port=8080)
