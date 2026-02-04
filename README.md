# Google Sheets to PDF Mail Merge Tool

A powerful, custom-built Google Apps Script tool that automates the process of merging spreadsheet data into Google Doc templates and converting them into professional PDF documents.

## 🚀 Features

- **Automated PDF Generation**: Merges rows from any selected worksheet into a master Google Doc template.
- **Dynamic File Naming**: Automatically names PDFs using the format `SheetName - CODIGO - AÑO`.
- **Automatic Sharing**: Sets generated PDF permissions to "Anyone with the link can view" automatically.
- **Smart Spreadsheet Updates**: Writes back the PDF ID, direct URL, a clickable link (using localized formula syntax), and a detailed status timestamp directly to your sheet.
- **Robust Mapping Engine**: Supports a JSON-based mapping system that links spreadsheet columns to `<<Placeholder>>` tags in your document.
- **Intelligent Error Handling**: The parser automatically corrects common copy-paste issues like "Smart/Curly Quotes" and provides descriptive feedback for configuration errors.
- **User-Friendly UI**: Includes a custom Google Sheets menu and a clean, Tailwind CSS-powered sidebar for selecting worksheets.

## 📋 Prerequisites

Your Google Spreadsheet must have a specific structure to function correctly.

### 1. The "Setup" Sheet
Create a tab named `Setup` with the following columns (Header in Row 1):
- **Column A (Worksheet Name)**: The exact name of the tab containing your data.
- **Column B (Master Doc ID)**: The ID of the Google Doc template (found in the URL).
- **Column C (Mapping)**: A JSON string mapping column numbers to placeholders.
  - *Example*: `{"2":"<<CODIGO>>", "5":"<<ESTUDIANTE>>", "6":"<<GRADO>>"}`
- **Column D (Folder ID)**: The ID of the Google Drive folder where PDFs should be saved.

### 2. The Data Worksheet(s)
Each worksheet you intend to merge must include the following:
- A column named exactly **CODIGO**.
- A column named exactly **AÑO**.
- Four columns for status tracking (the script will look for these keywords in the headers):
  - `Merged Doc ID`
  - `Merged Doc URL`
  - `Link to Merged Doc`
  - `Document Merge Status`

## 🛠️ Installation

1. Open your Google Spreadsheet.
2. Go to **Extensions** > **Apps Script**.
3. Create two files in the editor:
   - `Code.gs`: Paste the content from the `Code.gs` file in this repository.
   - `Sidebar.html`: Paste the content from the `Sidebar.html` file in this repository.
4. Refresh your spreadsheet. A new menu named **Merge Tools** will appear.

## 📖 How to Use

1. **Configure the Setup Sheet**: Fill in your template IDs, folder IDs, and mappings.
2. **Prepare your Data**: Ensure your data rows are ready. The script only processes rows where the `Merged Doc ID` column is empty (preventing duplicates).
3. **Run the Merge**:
   - Click **Merge Tools** > **Run Mail Merge**.
   - Select the desired worksheet from the dropdown in the sidebar.
   - Click **Run Merge**.
4. **Monitor Progress**: The status area will update you on the progress, and the spreadsheet will fill in with links to your generated PDFs in real-time.

## ⚠️ Troubleshooting Mapping JSON

The most common error is a malformed JSON string in the Mapping column. Ensure:
- All keys and values are wrapped in double quotes: `"key":"value"`.
- Each pair is separated by a comma.
- The entire string is wrapped in curly braces `{}`.

*Note: This script is configured to use the `;` separator for Google Sheets formulas (e.g., `=HYPERLINK(url; label)`), which is standard for many European and Latin American locales.*

## 📄 License
MIT License - Feel free to use and modify for your own projects!
