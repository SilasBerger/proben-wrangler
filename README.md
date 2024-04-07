# Proben-Wrangler
A collection of tools for wrangling Proben. 👨‍🏫

## General Usage
- `pip install -r requirements.txt`

## Tools
### Split Exams PDF
Splits a class-set PDF of exam scans into individual PDFs for each student, based on a points / grade list Excel spreadsheet.

**Usage**:
```shell
python split_exams_pdf.py <path_to_clas_pdf> <path_to_metadata_spreadsheet> <spreadsheet_tab_index>
```

**Example:** The following command splits a spreadsheet `Proben_xy_korrigiert.pdf` up into individual PDFs for each student, located in `/path/to/Proben_xy_korrigiert`, using student information in worksheet / tab `0` an Excel spreadsheet `Auswertung_Probe_xy.xlsx`.

```shell
python split_exams_pdf.py "/path/to/Proben_xy_korrigiert.pdf" "/Path/to/Auswertung_Probe_xy.xlsx" 0
```

**File requirements:**
- Column `Seite` contains the 1-indexed page number on which the exam for this student begins
- Column `Name` contains the student's last name
- Column `Vorname` contains the student's first name
- Column headers are located on line 3 (1-indexed)
- PDF contains no pages to be skipped (length of the printout is calculated as a different between to starting page values in the spreadsheet)
