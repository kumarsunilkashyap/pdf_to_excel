from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
import tabula
import os

app = Flask(__name__)

def extract_with_pdfplumber(file):
    dataframes = []
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if table:
                    cleaned = [row for row in table if any(cell.strip() if cell else '' for cell in row)]
                    if cleaned:
                        df = pd.DataFrame(cleaned[1:], columns=cleaned[0])
                        dataframes.append(df)
    return dataframes

def extract_with_tabula(file_path):
    try:
        dfs = tabula.read_pdf(file_path, pages='all', multiple_tables=True, lattice=True)
        return dfs
    except Exception as e:
        print("Tabula failed:", e)
        return []

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['pdf']
        if file and file.filename.endswith('.pdf'):
            temp_path = os.path.join("temp_" + file.filename)
            file.save(temp_path)

            # Try Tabula first
            dataframes = extract_with_tabula(temp_path)

            # If Tabula fails or returns nothing, fallback to pdfplumber
            if not dataframes:
                dataframes = extract_with_pdfplumber(temp_path)

            # If still nothing, fallback to plain text
            if not dataframes:
                with pdfplumber.open(temp_path) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if text:
                            lines = text.split('\n')
                            df = pd.DataFrame(lines, columns=[f'Page {page_num + 1} Text'])
                            dataframes.append(df)

            # Save to Excel
            output_path = 'output.xlsx'
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for idx, df in enumerate(dataframes):
                    df.to_excel(writer, sheet_name=f'Sheet{idx+1}', index=False)

            os.remove(temp_path)
            return send_file(output_path, as_attachment=True)
        else:
            return "कृपया एक वैध PDF फाइल अपलोड करें।"
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
