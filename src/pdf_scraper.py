import re
import PyPDF2
import pandas as pd


def extract_headers_and_content(pdf_path):
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)

        headers = []
        content = []

        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()

            # Try to extract the header (first line on the page)
            header_match = re.match(r'^(.*?)(?=\n|$)', page_text)
            if header_match:
                section_title = header_match.group(1)
            else:
                section_title = ""

            # Extract the content (including leading and trailing whitespace)
            section_text = page_text[len(section_title):]

            headers.append(section_title)
            content.append(section_text)

        return headers, content


def write_to_excel(headers, content, excel_path):
    # Create a pandas DataFrame with two columns: "Headers" and "Content"
    df = pd.DataFrame({'Headers': headers, 'Content': content})
    # Write the DataFrame to an Excel file
    df.to_excel(excel_path, index=False)
    print("Successfully written to Excel spreadsheet")


def main():
    pdf_path = '../pdf test samples/Bob was here.pdf'  # Replace with the path to your PDF file
    excel_path = '../Extraction file.xlsx'  # Replace with the desired output Excel file name

    headers, content = extract_headers_and_content(pdf_path)
    write_to_excel(headers, content, excel_path)


if __name__ == "__main__":
    main()
