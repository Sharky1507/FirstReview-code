# Text Correction Script Documentation

## Overview
This script is designed to process and correct text in Microsoft Word documents. It performs two main tasks:

1. **Word Replacement**: Replaces specific terms based on a predefined dictionary (`WORD_REPLACEMENTS`) while preserving the original case of the words.
2. **Name Formatting**: Formats names with initials and simplifies repeated mentions intelligently.

The script uses the `win32com.client` library to interact with Microsoft Word, allowing it to open, modify, and save Word documents programmatically.

The code is made to be modular.

---

## Approach

### 1. Word Replacement
The `replace_words` function scans the text for specific terms defined in the `WORD_REPLACEMENTS` dictionary. It replaces these terms while preserving their case (e.g., lowercase, uppercase, or title case).

### 2. Name Formatting
The `format_names` function identifies names with titles (e.g., Dr, Mr, Mrs) and formats them with initials. It also simplifies repeated mentions of names by using only the title and last name if the name has already been mentioned.

### 3. Processing Text Outside Quotes
The `process_text` function ensures that text inside quotes (e.g., dialogue) is not modified. It splits the text into segments based on quotes, processes the segments outside quotes, and then recombines them.

### 4. Word Document Processing
The `correct_document` function opens a Word document, processes its text using the above functions, and saves the corrected text to a new document.

---

## How to Use

### Prerequisites
1. **Python Environment**: Ensure you have Python installed on your system.
2. **Dependencies**: Install the required Python packages by running:
   ```bash
   pip install pywin32
   ```
   or by running:
  ```bash
   pip install -r requirements.txt
   ```

### Steps to Run
1. Place the input Word document in the same directory as the script and name it `input.docx`.
2. Run the script using the command:
   ```bash
   python app.py
   ```
3. The corrected document will be saved as `corrected.docx` in the same directory.

---
