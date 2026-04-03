import os
from pathlib import Path
import comtypes.client
import time

# Office PDF format constants
WD_FORMAT_PDF = 17
PP_FORMAT_PDF = 32

def convert_word_to_pdf(word_app, input_path, output_path):
    """Converts a Word document to PDF using Microsoft Word."""
    print(f"Processing Word: {input_path.name}")
    # Use absolute paths for COM
    abs_input = str(input_path.resolve())
    abs_output = str(output_path.resolve())
    
    doc = word_app.Documents.Open(abs_input)
    try:
        doc.SaveAs(abs_output, FileFormat=WD_FORMAT_PDF)
    finally:
        doc.Close()

def convert_powerpoint_to_pdf(ppt_app, input_path, output_path):
    """Converts a PowerPoint presentation to PDF using Microsoft PowerPoint."""
    print(f"Processing PowerPoint: {input_path.name}")
    # Use absolute paths for COM
    abs_input = str(input_path.resolve())
    abs_output = str(output_path.resolve())
    
    presentation = ppt_app.Presentations.Open(abs_input, WithWindow=False)
    try:
        presentation.SaveAs(abs_output, FileFormat=PP_FORMAT_PDF)
    finally:
        presentation.Close()

def main():
    current_dir = Path.cwd()
    
    # Initialize Word and PowerPoint applications
    word_app = None
    ppt_app = None
    
    try:
        print("Initializing Microsoft Word...")
        word_app = comtypes.client.CreateObject("Word.Application")
        word_app.Visible = False
        
        print("Initializing Microsoft PowerPoint...")
        ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
        
        # Identify files to convert
        files = [f for f in current_dir.iterdir() if f.suffix.lower() in [".doc", ".docx", ".pptx"]]
        
        if not files:
            print("No .doc, .docx, or .pptx files found in the current directory.")
            return

        for file_path in files:
            pdf_path = file_path.with_suffix(".pdf")
            
            # Skip if PDF already exists
            if pdf_path.exists():
                print(f"Skipping: {pdf_path.name} (already exists)")
                continue
                
            try:
                if file_path.suffix.lower() in [".doc", ".docx"]:
                    convert_word_to_pdf(word_app, file_path, pdf_path)
                elif file_path.suffix.lower() == ".pptx":
                    convert_powerpoint_to_pdf(ppt_app, file_path, pdf_path)
            except Exception as e:
                print(f"Failed to convert {file_path.name}: {e}")
                
    except Exception as e:
        print(f"An error occurred during initialization: {e}")
    finally:
        if word_app:
            try:
                word_app.Quit()
            except:
                pass
        if ppt_app:
            try:
                ppt_app.Quit()
            except:
                pass
        print("\nConversion process complete.")

if __name__ == "__main__":
    main()
