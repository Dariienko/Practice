# Code Explanation: word_take_info_to_exel.py

## Overview

This script automates the extraction of structured information from Word documents (`.docx`) containing meeting protocols. It processes both the text and tables in the documents, extracts relevant fields, and saves the results in an Excel file (`TOTAL_PROTOKOL.xlsx`). The script is designed to handle Ukrainian-language protocols but can be adapted for similar formats.

---

## Main Steps

1. **Read All Word Files**  
   The script scans the current directory for all `.docx` files and processes each one.

2. **Extract Text and Tables**  
   - `get_text(docx_path)`: Reads all paragraphs from the document and joins them into a single string.
   - `extract_tables(docx_path)`: Extracts all tables as lists of rows, where each row is a list of cell texts.

3. **Extract Meeting Info**  
   - `extract_meeting_info(text, tables)`:  
     - If tables are present, it extracts key fields (e.g., "Голова комітету", "Члени комітету") from the first table, collecting all relevant cell values below each keyword and joining them with commas.
     - If tables are not present, it falls back to extracting these fields from the text using regular expressions.

4. **Extract Questions (Agenda Items)**  
   - `extract_questions(text)`:  
     - Finds all agenda items ("Питання N:") and extracts the question text, speaker, votes, and decision for each.
     - Associates agenda topics with each question if available.

5. **Extract Voting Table**  
   - `extract_vote_table(text)`:  
     - Searches for a voting table at the end of the document and parses each member's votes.

6. **Extract Protocol Number and Date**  
   - `extract_doc_number_and_date(text)`:  
     - Uses regular expressions to extract the protocol number and meeting date.

7. **Build Output Rows**  
   - For each agenda item, creates a row with all extracted meeting info and the first member's voting data.
   - For each additional member, creates a row with only the member-specific data.
   - Appends the last table in the document as-is, mapping its columns by header.

8. **Save to Excel**  
   - All rows are combined into a pandas DataFrame and saved to `TOTAL_PROTOKOL.xlsx`.
   - The Excel columns are auto-sized for readability.

---

## Key Functions

- **get_text**: Reads all text from the document.
- **extract_tables**: Extracts all tables as lists of lists.
- **extract_meeting_info**: Extracts structured meeting info from the first table or from text.
- **extract_questions**: Extracts agenda items and related details.
- **extract_vote_table**: Extracts voting results from the text.
- **extract_doc_number_and_date**: Extracts protocol number and date.
- **Main block**: Orchestrates the extraction and saving process for all files.

---

## Output

- **TOTAL_PROTOKOL.xlsx**:  
  Contains all extracted data, including meeting info, agenda items, voting results, and the last table from each protocol file.

---

## Customization

- To adapt for other document formats or languages, adjust the keywords and regular expressions in the extraction functions.
- To change the output format, modify the DataFrame construction and Excel writing logic.
---