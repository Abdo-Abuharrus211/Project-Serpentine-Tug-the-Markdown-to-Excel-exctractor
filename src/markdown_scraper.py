import re
import pandas as pd
import os

# Define the maximum Excel cell string limit
MAX_CELL_STRING_LENGTH = 32767  # Adjust this value as needed

def extract_headings_and_content(markdown_file):
    # Try to open file, if it doesn't exist raise an error
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
                # Append the line to the current content
                current_content += line

        # Add the last heading and content
        if current_heading:
            headings.append(current_heading.strip('#').strip())
            heading_type.append(current_heading_type)
            content.append(current_content.strip())

        return headings, heading_type, content


def write_to_excel(headings, heading_type, content, excel_path):
    try:
        # Check if the Excel file exists
        if os.path.isfile(excel_path):
            # Append to the existing Excel file
            df = pd.read_excel(excel_path)
        else:
            # Create a new Excel file if it doesn't exist
            df = pd.DataFrame(columns=['Headers', 'Heading Type', 'Content'])

        # Initialize variables to keep track of remaining content
        remaining_content = ""
        current_heading = None

        for heading, h_type, section_content in zip(headings, heading_type, content):
            if current_heading != heading:
                # If the heading has changed, clear the remaining content
                remaining_content = ""
                current_heading = heading

            # Check if the content can fit within a single cell
            if len(section_content) <= MAX_CELL_STRING_LENGTH:
                df = df.append({'Headers': heading, 'Heading Type': h_type, 'Content': section_content},
                               ignore_index=True)
            else:
                # Split the content into multiple cells
                while section_content:
                    split_point = min(MAX_CELL_STRING_LENGTH, len(section_content))
                    cell_content = section_content[:split_point]
                    section_content = section_content[split_point:]
                    df = df.append({'Headers': heading, 'Heading Type': h_type, 'Content': cell_content},
                                   ignore_index=True)

            # Add any remaining content to the next cell
            if remaining_content:
                cell_content = remaining_content
                remaining_content = ""
                df = df.append({'Headers': heading, 'Heading Type': h_type, 'Content': cell_content}, ignore_index=True)

        # Write the DataFrame to the Excel file
        df.to_excel(excel_path, index=False)
        print(f"Written to Excel file '{excel_path}' successfully")
    except Exception as e:
        print(f"Error while writing to Excel file: {e}")


def main():
    # Replace with the path to your Markdown file
    markdown_file = '../markdown test samples/fab consolidated financial statement sample.md'
    # Replace with the desired output Excel file name
    excel_path = '../MD Extraction file.xlsx'

    try:
        headings, heading_type, content = extract_headings_and_content(markdown_file)
    except FileNotFoundError as e:
        print("File error:", e)
    else:
        write_to_excel(headings, heading_type, content, excel_path)


if __name__ == "__main__":
    main()
