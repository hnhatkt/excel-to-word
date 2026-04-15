from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from docxtpl import DocxTemplate
import os
import zipfile
import uuid
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# =====================
# HOME
# =====================
@app.route('/')
def index():
    return render_template('index.html')


# =====================
# GET SHEETS FROM EXCEL
# =====================
@app.route('/get-sheets', methods=['POST'])
def get_sheets():
    try:
        excel_file = request.files['excel']

        filename = secure_filename(excel_file.filename)
        temp_path = os.path.join(
            UPLOAD_FOLDER,
            f"{uuid.uuid4()}_{filename}"
        )

        excel_file.save(temp_path)

        xls = pd.ExcelFile(temp_path)

        return jsonify({
            "sheets": xls.sheet_names,
            "temp_path": temp_path
        })

    except Exception as e:
        return jsonify({"error": str(e)})


# =====================
# GENERATE WORD FILES
# =====================
@app.route('/generate', methods=['POST'])
def generate():
    try:
        sheet_name = request.form.get('sheet_name')
        temp_path = request.form.get('temp_path')
        template_file = request.files['template']

        if not temp_path or not os.path.exists(temp_path):
            return "❌ File Excel không tồn tại"

        df = pd.read_excel(temp_path, sheet_name=sheet_name)

        df = df.dropna(how='all')
        df = df.fillna('')

        if df.empty:
            return "❌ Không có dữ liệu"

        # save template
        template_name = secure_filename(template_file.filename)
        template_path = os.path.join(
            UPLOAD_FOLDER,
            f"{uuid.uuid4()}_{template_name}"
        )
        template_file.save(template_path)

        output_files = []

        for i, row in df.iterrows():
            context = row.to_dict()

            # clean NaN
            context = {k: ('' if pd.isna(v) else str(v)) for k, v in context.items()}

            # format date
            if 'ngay' in context:
                context['ngay'] = context['ngay'][:10]

            # =========================
            # FILE NAME RULE
            # so_hop_dong + goi_dich_vu + ten_san
            # =========================
            so_hop_dong = context.get('so_hop_dong', '')
            goi_dich_vu = context.get('goi_dich_vu', '')
            ten_san = context.get('ten_san', '')

            parts = []

            if so_hop_dong:
                parts.append(so_hop_dong)
            if goi_dich_vu:
                parts.append(goi_dich_vu)
            if ten_san:
                parts.append(ten_san)

            if parts:
                filename = "_".join(parts) + ".docx"
            else:
                filename = f"output_{i}.docx"

            filename = secure_filename(filename)

            output_path = os.path.join(OUTPUT_FOLDER, filename)

            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save(output_path)

            output_files.append(output_path)

        # =====================
        # ZIP RESULT
        # =====================
        zip_path = os.path.join(
            OUTPUT_FOLDER,
            f"{uuid.uuid4()}_result.zip"
        )

        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in output_files:
                zipf.write(file, os.path.basename(file))

        # =====================
        # CLEANUP FILES
        # =====================
        for file in output_files:
            if os.path.exists(file):
                os.remove(file)

        if os.path.exists(temp_path):
            os.remove(temp_path)

        if os.path.exists(template_path):
            os.remove(template_path)

        return send_file(
            zip_path,
            as_attachment=True,
            download_name="result.zip"
        )

    except Exception as e:
        return f"❌ Lỗi: {str(e)}"


# =====================
# RUN SERVER
# =====================
if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8000, debug=True)