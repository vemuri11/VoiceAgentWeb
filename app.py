from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
import io
import re
from datetime import datetime

app = Flask(__name__)

# In-memory data storage
data_rows = []

@app.route('/')
def home():
    return render_template('index.html', data=data_rows)

@app.route('/submit', methods=['POST'])
def submit():
    text = request.form.get('voiceText', '')
    name, place, paid, unpaid = parse_fields(text)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row = [name, place, paid, unpaid, timestamp]
    data_rows.append(row)
    return render_template('index.html', data=data_rows, message="Entry added!", voiceText='')

@app.route('/download_excel')
def download_excel():
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Place", "Paid", "Unpaid", "Timestamp"])
    for row in data_rows:
        ws.append(row)
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return send_file(file_stream, download_name="VoiceExcelData.xlsx", as_attachment=True)

def parse_fields(text):
    text = text.lower()
    name = re.search(r'name (\w+)', text)
    place = re.search(r'place (\w+)', text)
    paid = re.search(r'paid (\d+)', text)
    unpaid = re.search(r'unpaid (\d+)', text)
    return (
        name.group(1).capitalize() if name else '',
        place.group(1).capitalize() if place else '',
        paid.group(1) if paid else '',
        unpaid.group(1) if unpaid else ''
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000)
