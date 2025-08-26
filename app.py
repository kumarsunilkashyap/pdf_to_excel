from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['pdf']
        if file.filename.endswith('.pdf'):
            with pdfplumber.open(file) as pdf:
                all_tables = []
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            df = pd.DataFrame(table)
                            all_tables.append(df)
            output_path = 'output.xlsx'
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for idx, df in enumerate(all_tables):
                    df.to_excel(writer, sheet_name=f'Sheet{idx+1}', index=False)
            return send_file(output_path, as_attachment=True)
        else:
            return "कृपया एक वैध PDF फाइल अपलोड करें।"
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
