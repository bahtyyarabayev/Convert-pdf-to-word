import PyPDF2
from docx import Document

def convert_pdf_to_word(pdf_path, output_path):
    # Open the PDF file
    with open(pdf_path, 'rb') as pdf_file:
        # Create a PDF reader object
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        
        # Create a new Word document
        doc = Document()
        
        # Iterate over each page in the PDF
        for page_num in range(pdf_reader.numPages):
            # Extract the text from the page
            page = pdf_reader.getPage(page_num)
            text = page.extractText()
            
            # Add the extracted text to the Word document
            doc.add_paragraph(text)
        
        # Save the Word document
        doc.save(output_path)
        print(f"Conversion complete. Word file saved at {output_path}")

# Provide the path to the PDF file and the desired output path for the Word file
pdf_path = input("Enter the path to the PDF file: ")
output_path = input("Enter the output path for the Word file: ")

# Convert the PDF to Word
convert_pdf_to_word(pdf_path, output_path)
