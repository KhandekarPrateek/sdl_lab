from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import tabula
import PyPDF2

app = Flask(__name__)
UPLOAD_FOLDER = 'C:/Users/Prateek/OneDrive/Desktop/sdl_lab'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Function to extract paper code from the PDF text
def extract_paper_code(pdf_file_path):
    with open(pdf_file_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        full_text = ''
        for page in reader.pages:
            full_text += page.extract_text()

        # Search for the pattern "Paper Code" followed by the code
        # Adjust the search string according to how it appears in your PDFs
        search_keyword = "Paper Code"
        start_index = full_text.find(search_keyword)

        if start_index != -1:
            # Extract paper code (assuming it's right after the keyword)
            paper_code_start = start_index + len(search_keyword)
            paper_code = full_text[paper_code_start:].split()[0]  # Extract the next word after "Paper Code"
            return paper_code
        return None  # Return None if paper code not found

def pdf_to_dataframe(pdf_file_path):
    tables = tabula.read_pdf(pdf_file_path, pages='all')
    df = tables[0]

    # Ensure the required columns are present in the DataFrame
    df = df[['Enrollment No.', 'Student Name', 'Marks']]
    return df

def combine_pdfs_to_excel(pdf_files, excel_file_path):
    combined_df = None

    for pdf_file in pdf_files:
        df = pdf_to_dataframe(pdf_file)

        # Extract the paper code from the PDF
        paper_code = extract_paper_code(pdf_file)
        if not paper_code:
            paper_code = f'Subject_{pdf_files.index(pdf_file) + 1}'  # Default to a generic name if not found

        df = df[['Enrollment No.', 'Student Name', 'Marks']]
        df.columns = ['Enrollment No.', 'Student Name', paper_code]

        if combined_df is None:
            combined_df = df
        else:
            combined_df = pd.merge(combined_df, df, on=['Enrollment No.', 'Student Name'], how='outer')

    # Fill missing marks with 'Absent'
    combined_df.fillna('Absent', inplace=True)

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
