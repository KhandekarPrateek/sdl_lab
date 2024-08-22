import tabula
import pandas as pd

def pdf_to_dataframe(pdf_file_path):
    # Read PDF file and convert to dataframe
    tables = tabula.read_pdf(pdf_file_path, pages='all')
    # Assuming the first table is the required one
    df = tables[0]
    print(f"Columns in {pdf_file_path}: {df.columns}")  # Inspect the columns
    return df

def combine_pdfs_to_excel(pdf_file1, pdf_file2, excel_file_path):
    # Print file paths to debug
    print(f"PDF File 1: {pdf_file1}")
    print(f"PDF File 2: {pdf_file2}")

    # Convert both PDF files to dataframes
    df1 = pdf_to_dataframe(pdf_file1)
    df2 = pdf_to_dataframe(pdf_file2)

    # Inspect the columns before renaming
    print(f"Columns in df1 before renaming: {df1.columns}")
    print(f"Columns in df2 before renaming: {df2.columns}")

    # Select the relevant columns: 'Student Name' and 'Marks'
    df1 = df1[['Student Name', 'Marks']]
    df2 = df2[['Student Name', 'Marks']]

    # Rename columns to distinguish marks of different subjects
    df1.columns = ['Student Name', 'Subject1_Marks']
    df2.columns = ['Student Name', 'Subject2_Marks']

    # Merge dataframes on the 'Student Name' column
    combined_df = pd.merge(df1, df2, on='Student Name', how='outer')

    # Write the combined dataframe to an Excel file
    with pd.ExcelWriter(excel_file_path) as writer:
        combined_df.to_excel(writer, sheet_name='Combined Results', index=False)

# Example usage
pdf_file1 = '/home/khandekar/Desktop/excel/f2.pdf'
pdf_file2 = '/home/khandekar/Desktop/excel/f1.pdf'
excel_file_path = '/home/khandekar/Desktop/excel/combined_result.xlsx'

combine_pdfs_to_excel(pdf_file1, pdf_file2, excel_file_path)
