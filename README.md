# Batch File Conversion to PDF

This PowerShell script batch-converts various file types (PowerPoint, Word, Excel, and image files) in a folder and its subfolders into PDF format. The generated PDF files are saved in the same directory as the original files by default.

## Supported File Types
- **PowerPoint presentations** (`.ppt`, `.pptx`)
- **Word documents** (`.doc`, `.docx`)
- **Excel spreadsheets** (`.xls`, `.xlsx`)
- **Image files** (`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`)

## Prerequisites
1. **Windows PowerShell** with COM interop capabilities.
2. **Microsoft Office** installed, as the script uses Office's COM objects for conversion.
3. **Set Execution Policy:** Ensure that unsigned scripts can run on your system. To enable this, open PowerShell as an administrator and run:
   ```powershell
   Set-ExecutionPolicy Unrestricted
