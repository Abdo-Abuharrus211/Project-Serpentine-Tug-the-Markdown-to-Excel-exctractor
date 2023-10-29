import PyPDF2
import pandas as pd
from openpyxl import Workbook


# Function to extract text from PDF and return as a list of pages
def extract_text_from_pdf(pdf_path):
    pdf_text = []
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_text.append(page.extract_text())
    return pdf_text


# Function to write extracted text to Excel file
def write_to_excel(pdf_text, excel_path):
    # Create a pandas DataFrame with one row for each page's text
    columns = ['Page Text']
    df = pd.DataFrame(pdf_text, columns)
    # Write the DataFrame to an Excel file
    index = False
    df.to_excel(excel_path, index)


# Main function
def main():
    # Specify the path to the input PDF file and output Excel file
    pdf_path = '../pdf test samples/Bob was here.pdf'  # Replace 'input.pdf' with the path to your PDF file
    excel_path = '../Extraction file.xlsx'  # Replace 'output.xlsx' with the desired output Excel file name

    # Extract text from PDF
    pdf_text = extract_text_from_pdf(pdf_path)

    # Write extracted text to Excel
    write_to_excel(pdf_text, excel_path)
    print(f"Text extracted from {len(pdf_text)} pages and written to '{excel_path}'.")


# Run the main function if the script is executed
if __name__ == "__main__":
    main()
