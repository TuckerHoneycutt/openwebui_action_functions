# OpenWebUI Document Style Formatter

An OpenWebUI action function that extracts styling from uploaded DOCX or PDF documents and applies those styles to chat content, creating a professionally formatted document.

## Features

- **Extracts comprehensive styling** from source documents:
  - Fonts (name, size, color, bold, italic, underline)
  - Headers and footers
  - Tables with formatting
  - Page breaks and sections
  - Margins and indentations
  - Line spacing and paragraph alignment
  - First-line and hanging indents

- **Supports multiple formats**:
  - DOCX (Word documents)
  - PDF (converted to DOCX for style extraction)

- **Applies styles to chat content**:
  - Formats chat messages with extracted styles
  - Preserves document structure (headers, footers, tables)
  - Maintains professional formatting

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

2. Place `main.py` in your OpenWebUI functions directory

3. Restart OpenWebUI or reload functions

## Usage

1. **After a chat session**, click the action button that appears
2. **Upload a DOCX or PDF document** that contains the styling you want to replicate
3. The function will:
   - Extract all styling information from your document
   - Format the chat content using those styles
   - Generate a new DOCX file with the formatted content
4. **Download** the formatted document

## How It Works

1. **Style Extraction**: The function analyzes the uploaded document and extracts:
   - Font properties (name, size, color, formatting)
   - Paragraph formatting (alignment, spacing, indentation)
   - Document structure (headers, footers, sections)
   - Table formatting
   - Page layout (margins, page breaks)

2. **Style Application**: The extracted styles are applied to the chat messages:
   - Each message is formatted according to the document's paragraph styles
   - Role labels (User/Assistant) are styled appropriately
   - Headers and footers from the source document are added
   - Tables are preserved if present

3. **Output Generation**: A new DOCX document is created with:
   - All chat content formatted with extracted styles
   - Original document structure preserved
   - Professional formatting maintained

## Requirements

- Python 3.7+
- python-docx
- PyMuPDF (fitz)
- pdf2docx

## Notes

- The function works best with well-formatted source documents
- Complex formatting may require manual adjustments
- PDF files are converted to DOCX internally for style extraction
- The function automatically handles file cleanup

## Troubleshooting

- **"No file provided"**: Make sure to upload a document when prompted
- **"No chat messages found"**: Ensure you're using the function after a chat session with messages
- **Style extraction issues**: Try with a simpler document first to test functionality
- **PDF conversion errors**: Ensure the PDF is not password-protected or corrupted
