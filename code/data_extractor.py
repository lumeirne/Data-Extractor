import re
import os
import pandas as pd
import PyPDF2
from docx import Document

# Main function to extract data from files in input directory and save to Excel in output directory
def data_extractor_main(input_directory, output_directory):
    # Function to extract emails from text using regular expression
    def extract_emails(text):
        return re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)

    # Function to extract phone numbers from text using regular expression
    def extract_phone_numbers(text):
        return re.findall(r'\b(?:\d{3}[-.\s]?)?\d{3}[-.\s]?\d{4}\b', text)

    # Function to extract text from PDF file
    def extract_from_pdf(file_path):
        pdf_text = ""
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                pdf_text += page.extract_text()
        return pdf_text
    
    # Function to extract text from DOCX file
    def extract_from_docx(file_path):
        doc = Document(file_path)
        docx_text = ""
        for paragraph in doc.paragraphs:
            docx_text += paragraph.text + "\n"
        return docx_text
    
    # Function to extract text from DOC file
    def extract_from_doc(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
           doc_text = f.read()
        return doc_text

    # Dictionary to store extracted data
    data = {'File Name': [], 'Email': [], 'Phone Number': [], 'Text': []}

    # Loop through files in input directory
    for filename in os.listdir(input_directory):
        file = os.path.join(input_directory, filename)
        # Extract text based on file type
        if filename.endswith('.pdf'):
            text = extract_from_pdf(file)
        elif filename.endswith('.docx'):
            text = extract_from_docx(file)
        elif filename.endswith('.doc'):
            text = extract_from_doc(file)
        else:
            continue

        # Extract emails and phone numbers from text
        emails = extract_emails(text)
        phones = extract_phone_numbers(text)

        # Remove duplicates
        emails = list(set(emails))
        phones = list(set(phones))

        # Add data to dictionary
        data['File Name'].append(filename)
        data['Email'].append(emails)
        data['Phone Number'].append(phones)
        data['Text'].append(text)

    # Path for output Excel file
    output_path = os.path.join(output_directory, 'Output.xls')
    # Convert dictionary to DataFrame and save to Excel
    extractor_df = pd.DataFrame(data)
    #To save output in .xls format we need to use engine openpyxl
    extractor_df.to_excel(output_path, index=False, engine='openpyxl')

# Entry point of the script
if __name__ == "__main__":
    # Directories for input and output
    dir_path = os.path.dirname(os.getcwd())
    input_directory = os.path.join(dir_path, 'input')
    output_directory = os.path.join(dir_path, 'output')
    # Call main function to extract data
    data_extractor_main(input_directory, output_directory)
