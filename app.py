from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import tabula

app = Flask(__name__)
UPLOAD_FOLDER = '/home/khandekar/Desktop/excel/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def pdf_to_dataframe(pdf_file_path):
    tables = tabula.read_pdf(pdf_file_path, pages='all')
    df = tables[0]
    return df

def combine_pdfs_to_excel(pdf_files, excel_file_path):
    combined_df = None

    for i, pdf_file in enumerate(pdf_files):
        df = pdf_to_dataframe(pdf_file)
        df = df[['Student Name', 'Marks']]
        df.columns = ['Student Name', f'Subject{i+1}_Marks']

        if combined_df is None:
            combined_df = df
        else:
            combined_df = pd.merge(combined_df, df, on='Student Name', how='outer')

    combined_df.to_excel(excel_file_path, index=False)

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if request.method == 'POST':
        pdf_files = request.files.getlist('pdfs')

        # Save each PDF file
        pdf_paths = []
        for pdf_file in pdf_files:
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
            pdf_file.save(pdf_path)
            pdf_paths.append(pdf_path)

        # Combine and convert to Excel
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'combined_result.xlsx')
        combine_pdfs_to_excel(pdf_paths, excel_file_path)

        return send_file(excel_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
