import re
import pandas as pd
import openpyxl
import os

MAX_CELL_STRING_LENGTH = 32767
"""
Extract headings and content from a Markdown file and write to an Excel file.

:param markdown_file: a string
:precondition: markdown_file is a valid path to a Markdown file
:postcondition: extract the headings and content from the Markdown file
:return: a tuple containing a list of headings, a list of heading types, and a list of content
"""


def extract_headings_and_content(markdown_file):
    try:
        if os.path.isfile(markdown_file):
            print(f"The file '{markdown_file}' exists.")
        else:
            raise FileNotFoundError("File not found in directory")
    finally:
        with open(markdown_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        headings = []
        heading_type = []
        content = []

        current_heading = ""
        current_content = ""
        current_heading_type = 0
        for line in lines:
            # Print the line for debugging
            print("Processing line:", line)

            # Check if the line is a heading (starts with one or more '#' symbols)
            match = re.match(r'^(#+)\s*', line)
            if match:
                # Store the previous heading and content
                if current_heading:
                    headings.append(current_heading.strip('#').strip())
                    heading_type.append(current_heading_type)
                    content.append(current_content.strip())
                # Update the current heading
                current_heading = line
                current_heading_type = len(match.group(1))
                current_content = ""
            else:
                current_content += line

        # Add the last heading and content
        if current_heading:
            headings.append(current_heading.strip('#').strip())
            heading_type.append(current_heading_type)
            content.append(current_content.strip())

        return headings, heading_type, content


"""
Write the headings and content to an Excel file.

:param headings: a list of strings
:param heading_type: a list of integers
:param content: a list of strings
:param excel_path: a string
:precondition: headings, heading_type, and content are the same length
:precondition: excel_path is a valid path to an Excel file
:postcondition: write the headings and content to the Excel file
:return: None

:raises FileNotFoundError: if the Excel file does not exist
:raises Exception: if there is an error while writing to the Excel file
"""


def write_to_excel(headings, heading_type, content, excel_path):
    try:
        # Check if the Excel file exists
        if os.path.isfile(excel_path):
            # Append to the existing Excel file
            df = pd.read_excel(excel_path)
        else:
            # Create a new Excel file if it doesn't exist
            df = pd.DataFrame(columns=['Headers', 'Heading Type', 'Content'])

        remaining_content = ""
        current_heading = None
        data_to_append = []

        for heading, h_type, section_content in zip(headings, heading_type, content):
            if current_heading != heading:
                # If the heading has changed, clear the remaining content
                remaining_content = ""
                current_heading = heading

            # Check if the content can fit within a single cell
            if len(section_content) <= MAX_CELL_STRING_LENGTH:
                data_to_append.append({'Headers': heading, 'Heading Type': h_type, 'Content': section_content})
            else:
                # Split the content into multiple cells
                while section_content:
                    split_point = min(MAX_CELL_STRING_LENGTH, len(section_content))
                    cell_content = section_content[:split_point]
                    section_content = section_content[split_point:]
                    data_to_append.append({'Headers': heading, 'Heading Type': h_type, 'Content': cell_content})

            # Add any remaining content to the next cell
            if remaining_content:
                cell_content = remaining_content
                remaining_content = ""
                data_to_append.append({'Headers': heading, 'Heading Type': h_type, 'Content': cell_content})

        # Create a new DataFrame with the data to append
        new_data = pd.DataFrame(data_to_append)

        # Concatenate the existing DataFrame with the new data
        df = pd.concat([df, new_data], ignore_index=True)

        # Write the DataFrame to the Excel file
        df.to_excel(excel_path, index=False)
        print(f"Written to Excel file '{excel_path}' successfully")
    except Exception as e:
        print(f"Error while writing to Excel file: {e}")


def main():
    markdown_file = input("Enter the path to the Markdown file: ");
    # Replace with the desired output Excel file name
    excel_path = input("Please enter the path to the Excel file you want to write to:")
    try:
        headings, heading_type, content = extract_headings_and_content(markdown_file)
    except FileNotFoundError as e:
        print("File error:", e)
    else:
        write_to_excel(headings, heading_type, content, excel_path)
    finally:
        print("Done")


if __name__ == "__main__":
    main()
