import win32clipboard
import win32com.client
from ollama import chat
from ollama import ChatResponse
import sys
import re  # Import regex to clean file names
import keyboard  # Library to detect keypress

# Define the Word document path
word_file = r"C:\Users\RGMatr1x\Downloads\Job_Application_Files\CoverLetter\Cover_Letter_with_Header.docx"
MODEL = 'qwen2.5:0.5b'


def edit_pdf_file_name(company_name):
    """Generate a clean, Windows-safe filename for the PDF."""
    # Remove any illegal characters for Windows filenames
    safe_company_name = re.sub(r'[<>:"/\\|?*]', '', company_name).strip()
    
    # Ensure company name is not empty after cleaning
    if not safe_company_name:
        safe_company_name = "Unknown_Company"

    pdf_file = rf"C:\Users\RGMatr1x\Downloads\Job_Application_Files\CoverLetter\{safe_company_name}_CoverLetter.pdf"
    return pdf_file


def get_clipboard_text() -> str:
    """Retrieve text content from the clipboard."""
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            return win32clipboard.GetClipboardData().strip()
        else:
            return ""
    finally:
        win32clipboard.CloseClipboard()


def get_company_name(cover_letter: str) -> str:
    """
    Extract the company name from the cover letter using AI.
    
    Returns:
        - str: Extracted company name (cleaned and formatted for file naming).
    """
    response: ChatResponse = chat(model=MODEL, messages=[{
        'role': 'user',
        'content': f"""
        Extract and return ONLY the company name from the following cover letter.
        Do NOT include any additional text or formatting.
        If no company name is explicitly mentioned, return "Unknown Company" only.

        Cover Letter:
        {cover_letter}
        """
    }])

    # Extract AI response and remove any unnecessary spaces or special characters
    company_name = response.message.content.strip()
    
    # Print raw response for debugging
    print(f"Raw Company Name: {company_name}")

    # Clean the extracted company name to be filename-safe
    company_name = re.sub(r'[<>:"/\\|?*]', '', company_name).strip()

    # If the cleaned name is empty or too short, default to "Unknown Company"
    if not company_name or len(company_name) < 2:
        company_name = "Unknown_Company"

    print(f"Final Processed Company Name: {company_name}")
    return company_name


def insert_text_into_word(file_path, company_name):
    """Open a Word document, clear existing content, paste clipboard text, set font to Aptos (Body), justify alignment, and save as PDF."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run Word in the background
    
    # Open the existing document
    doc = word.Documents.Open(file_path)
    
    # Select all and delete existing content
    selection = word.Selection
    selection.WholeStory()  # Select all text
    selection.TypeBackspace()  # Delete everything

    # Paste clipboard content
    selection.Paste()

    # Select all text again to apply formatting
    selection.WholeStory()

    # Set font to "Aptos" (Body)
    selection.Font.Name = "Aptos"

    # Justify alignment
    selection.ParagraphFormat.Alignment = 3  # 3 = Justify

    # Save the document
    doc.Save()

    # Export to PDF with a cleaned filename
    pdf_file = edit_pdf_file_name(company_name)
    doc.ExportAsFixedFormat(pdf_file, 17)  # 17 = wdExportFormatPDF

    # Close Word
    doc.Close(SaveChanges=True)
    word.Quit()

    print(f"Saved PDF: {pdf_file}")


def main():
    """Main function that runs when F9 is pressed."""
    print("Program is running... Press 'F9' to generate a cover letter.")
    
    while True:
        keyboard.wait("`")  # Wait for ` key press
        print("\n ` detected! Processing cover letter...\n")

        cover_letter_from_gpt = get_clipboard_text()
        
        if not cover_letter_from_gpt:
            print("Clipboard is empty or does not contain text.")
            continue  # Go back to waiting for the next F9 press
        
        company_name = get_company_name(cover_letter_from_gpt)
        insert_text_into_word(word_file, company_name)

        print("\nCover letter saved. Press 'F9' again to generate a new one.\n")


if __name__ == "__main__":
    main()
