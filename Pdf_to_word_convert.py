from pdf2docx import Converter
import io
import os


def pdf_to_word(file_stream, output_path):
    # Wrap the file stream as BytesIO for pdf2docx
    pdf_io = io.BytesIO(file_stream.read())
    cv = Converter(pdf_io)

    # Convert to Word and save to output path
    cv.convert(output_path, start=0, end=None)
    cv.close()
    print(f"Conversion complete! PDF has been converted to '{output_path}'.")


# Define file paths for local testing
pdf_path = os.path.expanduser('~/Downloads/Payoff Documents.pdf')  # Your PDF file location
output_path = os.path.expanduser('~/Downloads/Payoff_Documents.docx')  # Output location for the Word file

# Open the PDF file in binary mode and test the function
with open(pdf_path, 'rb') as pdf_file:
    pdf_to_word(pdf_file, output_path)
