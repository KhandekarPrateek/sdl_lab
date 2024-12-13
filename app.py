from flask import Flask, render_template, request, send_file, jsonify
import os
import pandas as pd
import tabula
import PyPDF2

app = Flask(__name__)
UPLOAD_FOLDER = 'C:/Users/Prateek/OneDrive/Desktop/sdl_lab'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_paper_code(pdf_file_path):
    with open(pdf_file_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        full_text = ''
        for page in reader.pages:
            full_text += page.extract_text()

        search_keyword = "Paper Code"
        start_index = full_text.find(search_keyword)

        if start_index != -1:
            paper_code_start = start_index + len(search_keyword)
            paper_code = full_text[paper_code_start:].split()[0]
            return paper_code
        return None

def pdf_to_dataframe(pdf_file_path):
    tables = tabula.read_pdf(pdf_file_path, pages='all')
    df = tables[0]
    df = df[['Enrollment No.', 'Student Name', 'Marks']]
    return df

def combine_pdfs_to_dataframe(pdf_files):
    combined_df = None

    for pdf_file in pdf_files:
        df = pdf_to_dataframe(pdf_file)
        paper_code = extract_paper_code(pdf_file)
        if not paper_code:
            paper_code = f'Subject_{pdf_files.index(pdf_file) + 1}'

        df.columns = ['Enrollment No.', 'Student Name', paper_code]

        if combined_df is None:
            combined_df = df
        else:
            combined_df = pd.merge(combined_df, df, on=['Enrollment No.', 'Student Name'], how='outer')

    combined_df.fillna('Absent', inplace=True)
    return combined_df

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/merge', methods=['POST'])
def merge_files():
    if request.method == 'POST':
        pdf_files = request.files.getlist('pdfs')

        pdf_paths = []
        for pdf_file in pdf_files:
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
            pdf_file.save(pdf_path)
            pdf_paths.append(pdf_path)

        combined_df = combine_pdfs_to_dataframe(pdf_paths)

        # Convert DataFrame to HTML for display
        merged_data_html = combined_df.to_html(classes='table table-striped', index=False)

        # Save Excel file to the server (for downloading later)
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'combined_result.xlsx')
        combined_df.to_excel(excel_file_path, index=False)

        return jsonify({'merged_data': merged_data_html})

@app.route('/download', methods=['GET'])
def download_file():
    excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'combined_result.xlsx')
    return send_file(excel_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
