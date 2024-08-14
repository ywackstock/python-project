from flask import Flask, request, jsonify
import pandas as pd
import os
from fpdf import FPDF
import matplotlib.pyplot as plt

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'there is no part in the file'}), 400
    file = request.files['9']
    if file.filename == '':
        return jsonify({'error': 'no file selected'}), 400
    if file and file.filename.endswith('.xlsx'):
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        xls = pd.ExcelFile(file_path)
        return jsonify({
            'file_path': file_path,
            'names_of_sheet': xls.sheet_names,
            'num_of_sheets': len(xls.sheet_names)
        })
    else:
        return jsonify({'error': 'file not valid'}), 400

@app.route('/process', methods=['POST'])
def data_processing():
    data = request.json
    file_path = data['file_path']
    operations = data['operations']
    report = {}
    xls = pd.ExcelFile(file_path)
    for sheet in operations:
        sheet_name = sheet['sheet_name']
        type_operation = sheet['operation']
        columns = sheet['columns']
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if type_operation == 'sum':
            result = df[columns].sum().to_dict()
        elif type_operation == 'average':
            result = df[columns].mean().to_dict()
            report[sheet_name] = result
    return jsonify(report)

@app.route('/create_pdf', methods=['POST'])
def create_pdf():
    report = request.json
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Calibri', 'B', 16)
    for name_of_sheet, data in report.items():
        pdf.cell(200, 10, txt=name_of_sheet, ln=True)
        for col, value in data.items():
            pdf.cell(200, 10, txt=f"{col}: {value}", ln=True)
    pdf_path = 'report.pdf'
    pdf.output(pdf_path)
    return jsonify({'pdf_path': pdf_path})

@app.route('/grafh', methods=['POST'])
def plot_graph():
    data = request.json
    sums = {sheet: sum(val.values()) for sheet, val in data.items()}
    sheets = list(sums.keys())
    values = list(sums.values())
    plt.bar(sheets, values)
    plt.xlabel('names sheet')
    plt.ylabel('sum')
    plt.title('amount of each sheet')
    plt.savefig('sums_of_sheets.png')
    return jsonify({'graph_pat': 'sums_of_sheets.png'})

@app.route('/detailed_pdf', methods=['POST'])
def detailed_pdf():
    report = request.json
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Calibri', 'B', 16)
    pdf.cell(200, 10, txt="details of the report", ln=True)
    for name_of_sheet, data in report.items():
        pdf.cell(200, 10, txt=name_of_sheet, ln=True)
        for col, value in data.items():
            pdf.cell(200, 10, txt=f"{col}: {value}", ln=True)
    pdf.add_page()
    pdf.cell(200, 10, txt="graphs", ln=True)
    pdf.image('sheet_sums.png', x=10, y=30, w=180)
    pdf_path = 'detailed_report.pdf'
    pdf.output(pdf_path)
    return jsonify({'pdf_path': pdf_path})

if __name__ == '__main__':
    app.run(debug=True)