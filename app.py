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

def combine_pdfs_to_excel(pdf_file1, pdf_file2, excel_file_path):
    df1 = pdf_to_dataframe(pdf_file1)
    df2 = pdf_to_dataframe(pdf_file2)

    df1 = df1[['Student Name', 'Marks']]
    df2 = df2[['Student Name', 'Marks']]

    df1.columns = ['Student Name', 'Subject1_Marks']
    df2.columns = ['Student Name', 'Subject2_Marks']

    combined_df = pd.merge(df1, df2, on='Student Name', how='outer')

    combined_df.to_excel(excel_file_path, index=False)

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if request.method == 'POST':
        pdf1 = request.files['pdf1']
        pdf2 = request.files['pdf2']
        
        pdf1_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf1.filename)
        pdf2_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf2.filename)
        
        pdf1.save(pdf1_path)
        pdf2.save(pdf2_path)
        
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'combined_result.xlsx')
        combine_pdfs_to_excel(pdf1_path, pdf2_path, excel_file_path)
        
        return send_file(excel_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
