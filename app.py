from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from docxtpl import DocxTemplate
import os
import zipfile

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')


# ===== LẤY DANH SÁCH SHEET =====
@app.route('/get-sheets', methods=['POST'])
def get_sheets():
    excel_file = request.files['excel']

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(excel_path)

    xls = pd.ExcelFile(excel_path)
    sheets = xls.sheet_names

    return jsonify({
        "sheets": sheets,
        "path": excel_path
    })


# ===== GENERATE FILE =====
@app.route('/generate', methods=['POST'])
def generate():
    try:
        excel_path = request.form['excel_path']
        sheet_name = request.form.get('sheet_name')
        word_template = request.files['template']

        template_path = os.path.join(UPLOAD_FOLDER, word_template.filename)
        word_template.save(template_path)

        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        df = df.dropna(how='all')
        df = df.fillna('')

        if df.empty:
            return "❌ File Excel không có dữ liệu"

        output_files = []

        for i, row in df.iterrows():
            context = row.to_dict()

            if 'ngay' in context:
                context['ngay'] = str(context['ngay'])[:10]

            doc = DocxTemplate(template_path)
            doc.render(context)

            filename = f"output_{i}.docx"
            if 'ten' in context and context['ten']:
                filename = f"{context['ten']}.docx"

            output_path = os.path.join(OUTPUT_FOLDER, filename)
            doc.save(output_path)

            output_files.append(output_path)

        zip_path = os.path.join(OUTPUT_FOLDER, 'result.zip')

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in output_files:
                zipf.write(file, os.path.basename(file))

        return send_file(zip_path, as_attachment=True)

    except Exception as e:
        return f"❌ Lỗi: {str(e)}"


if __name__ == '__main__':
    app.run(debug=True)