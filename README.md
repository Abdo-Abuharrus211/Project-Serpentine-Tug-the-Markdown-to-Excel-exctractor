# Project Title: Serpentine Tug

## Objective: Extract Heading and content from Markdown files.

## Author: Abdulqadir Abuharrus <br> Date: 29/10/2023 <br>Version: 1.2

___

## Purpose

Extracting all header and their content (section by section) from a PDF file and inserting it into an
Excel spreadsheet with the purpose of feeding the data to a Large Language Model (LLM).<br>
The target file's format was a PDF file that is 110+ pages long and thus wanted to write a program to complete the task.

## My Role

I received a request from a friend that needed help with this task as they're less experienced programmers,
and I took up the challenge as little fun project on the side.

## Requirements/Goals

1. To parse through the file, identify each section by its heading/subheading.
2. For each section extract its heading and insert it into a "Headings" column in an Excel spreadsheet.
3. For each section, extract the content and insert it into a "Content" column in the same Excel spreadsheet.
4. The format of the Excel file is to have each heading in a cell with the cell to the right containing its content.
5. We need to preserve all white space, punctuation and other string formatting preexisting in the text, for the LLM to
   understand the data.
6. The formatting of content within cells is irrelevant and has no effect on the quality of the input for the LLM, and
   therefore the content can be <br> inserted as is into its designated cell without formatting,
   provided it's not over the limit.

## Remarks

* After a few trial runs and iterations, I found that parsing for different headings and subheadings in a PDF file to be
  challenging and inconsistent.
* A query regarding converting the PDF files to markdown files has been submitted, as they're far easier to parse
  through
  and yield consistent results thanks to Markdown's formatting.
* If a clean 1:1 conversion is possible for such files, then that might be a better route to pursue. Awaiting
  confirmation...
* After getting the confirmation that the PDF files can be converted to Markdown files, I've decided to pursue that
  route instead as it's far easier to parse through and yields consistent results.
* The PDF files can be reliably and securely converted to Markdown files using a variety of tools on the web.
* With the files converted, the script parses through the Markdown files and extracts the data as intended.
* However, the script is unable to extract data from graphs, tables, and illustrations, and thus the data is lost.
* The script is able to extract data from the Markdown files and insert it into an Excel spreadsheet as intended.
* I will attempt to find a solution to the issue of overlooked data in graphs, tables, and illustrations...

## Use instructions

_**To be updated once complete**_

#### Todo:

- [x] ~~Create a function to extract all headings and their content from a PDF file.~~
- [x] ~~Create a function to insert the extracted data into an Excel spreadsheet.~~
- [x] ~~Create a function to extract all headings and their content from a Markdown file.~~
- [x] ~~Create a function to insert the extracted data into an Excel spreadsheet.~~
- [ ] Tend to the issue of overlooked data in graphs, tables, and illustrations. 