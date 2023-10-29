import re
import pandas as pd


def extract_headings_and_content(markdown_file):
    with open(markdown_file, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    headings = []
    heading_type = []
    content = []

    current_heading = ""
    current_content = ""
    current_heading_type = 0  # Default heading type
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

    # Debugging output
    # print("Length of headings:", len(headings))
    # print("Length of heading_type:", len(heading_type))
    # print("Length of content:", len(content))

    return headings, heading_type, content


def write_to_excel(headings, heading_type, content, excel_path):
    # Create a pandas DataFrame with three columns: "Headers", "Heading Type", and "Content"
    df = pd.DataFrame({'Headers': headings, 'Heading Type': heading_type, 'Content': content})
    # Write the DataFrame to an Excel file
    df.to_excel(excel_path, index=False)


def main():
    markdown_file = 'fab conoslidated financial statement sample.md'  # Replace with the path to your Markdown file
    excel_path = 'MD Extraction file.xlsx'  # Replace with the desired output Excel file name

    headings, heading_type, content = extract_headings_and_content(markdown_file)
    write_to_excel(headings, heading_type, content, excel_path)


if __name__ == "__main__":
    main()
