from PyPDF2 import PdfMerger
from PIL import Image
import os
from tkinter import Tk, filedialog
import win32com.client  # Requires installation of pywin32 for Word doc conversion
import time

def image_to_pdf(image_path):
    """Convert an image to a single-page PDF"""
    try:
        image = Image.open(image_path)
        pdf_path = image_path.replace(image_path.split('.')[-1], 'pdf')  # Replace extension with .pdf
        image.save(pdf_path, "PDF", resolution=100.0)
        return pdf_path
    except Exception as e:
        print(f"Error converting image {image_path} to PDF: {e}")
        return None

def word_to_pdf(docx_path):
    """Convert a Word document (.docx) to PDF using Microsoft Word (Windows only)"""
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(docx_path)
        pdf_path = docx_path.replace(".docx", ".pdf")  # Change extension to .pdf
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the code for PDF format
        doc.Close()
        word.Quit()
        return pdf_path
    except Exception as e:
        print(f"Error converting Word document {docx_path} to PDF: {e}")
        return None

def wait_for_file(file_path):
    """Wait for the file to be accessible, checking every 500ms."""
    while True:
        try:
            with open(file_path, 'rb'):
                break  # File is available
        except IOError:
            time.sleep(0.5)  # Wait a bit before trying again

def is_file_locked(file_path):
    """Check if the file is locked by another process"""
    try:
        with open(file_path, 'a+b') as f:
            pass  # Try to open the file
        return False
    except IOError:
        return True

def remove_file_after_delay(file_path, delay=1, max_retries=10):
    """Attempt to remove the file after a delay if it is locked, retry up to max_retries times"""
    retries = 0
    while retries < max_retries:
        if not is_file_locked(file_path):
            try:
                os.remove(file_path)
                print(f"Removed file: {file_path}")
                return
            except PermissionError:
                print(f"Permission error removing {file_path}. Retrying...")
                time.sleep(delay)
                retries += 1
        else:
            print(f"File {file_path} is locked. Waiting for release...")
            time.sleep(delay)
            retries += 1
    
    print(f"Failed to remove {file_path} after {max_retries} retries.")

def merge_pdfs(output_filename="merged_document.pdf"):
    # Create Tkinter root window (hidden)
    root = Tk()
    root.withdraw()
    root.title("PDF Merger")
    
    # Ask the user to select multiple PDF, Word, and Image files
    file_paths = filedialog.askopenfilenames(
        title="Select PDF, Word, or Image Files to Merge",
        filetypes=[("PDF, Word, and Image Files", "*.pdf;*.docx;*.jpg;*.jpeg;*.png")]
    )
    
    if not file_paths:
        print("No files selected. Exiting.")
        return
    
    # Create a PdfMerger instance
    merger = PdfMerger()
    
    # Ensure the files are processed in the exact order they were selected
    for file_path in file_paths:
        if file_path.lower().endswith(('jpg', 'jpeg', 'png')):
            # Convert image to PDF and append it to the merger
            print(f"Converting image {file_path} to PDF...")
            pdf_file = image_to_pdf(file_path)
            if pdf_file:
                wait_for_file(pdf_file)  # Ensure the file is fully accessible before adding
                merger.append(pdf_file)
                remove_file_after_delay(pdf_file)  # Remove the temporary image PDF after merging
        elif file_path.lower().endswith('docx'):
            # Convert Word document to PDF and append it to the merger
            print(f"Converting Word document {file_path} to PDF...")
            pdf_file = word_to_pdf(file_path)
            if pdf_file:
                wait_for_file(pdf_file)  # Ensure the file is fully accessible before adding
                merger.append(pdf_file)
                remove_file_after_delay(pdf_file)  # Remove the temporary Word PDF after merging
        else:
            # Append PDF directly
            print(f"Appending PDF {file_path}...")
            merger.append(file_path)
    
    # Save the merged PDF to the specified output file
    output_path = filedialog.asksaveasfilename(
        title="Save Merged PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        initialfile=output_filename
    )
    
    if output_path:
        try:
            merger.write(output_path)
            print(f"Merged PDF saved to {output_path}")
        except Exception as e:
            print(f"Error saving merged PDF: {e}")
        finally:
            # Ensure that the merger is closed after processing
            merger.close()
            print("PDF merger closed successfully.")
    else:
        print("Save operation canceled.")

if __name__ == "__main__":
    merge_pdfs()
