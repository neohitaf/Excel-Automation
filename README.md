# Student Payroll Excel Automation

A Python-based workflow automation project for processing part-time student employment records and generating institution-compatible Excel outputs.

The project was developed to reduce repetitive manual data entry in a university administrative workflow while preserving the institution's existing Excel templates, formulas, formatting, and manual review process.

## Project Context

The existing workflow relies on multiple Word and Excel documents:

* **Form-3 Word documents** containing student identity, banking, and GSS information
* **Timesheet Excel files** containing unit, role, and total working-hour information
* Institution-specific **payroll, tax declaration, and bank-list templates**

Manually transferring this information between files is repetitive and vulnerable to data-entry errors. This project automates the extraction, matching, calculation, and placement steps without replacing the Excel-based control process already used by administrative staff.

## Design Approach

This project intentionally follows a **template-preserving and human-in-the-loop** approach.

The output workbooks are not recreated from scratch. Instead, the application opens the institution's existing Excel templates and writes data only to the designated input cells.

This decision preserves:

* Existing workbook layouts and formatting
* Institution-defined formulas
* Formula cells that must remain visible even when their result is `0`
* Fixed sections used for different employee document types
* The ability for authorized staff to review and manually adjust the final workbook

The fixed row positions in the code therefore represent the structure of the approved templates rather than a generic dynamic spreadsheet design.

## Current Features

* Parses Form-3 `.docx` files and locates the relevant student table
* Supports alternative column-title variations in source documents
* Extracts student identity, GSS, bank, IBAN, and related information
* Reads `.xlsx` timesheet files and identifies student records
* Extracts unit names, payroll periods, and total working hours
* Matches Form-3 records with timesheet records
* Calculates premium days, missing days, and payroll-related fields
* Separates records according to GSS status and document type
* Preserves the original formulas and formatting in institutional templates
* Generates the following Excel outputs:

  * `bordro.xlsx`
  * `muhtasar.xlsx`
  * `banka_listesi.xlsx`
* Reports missing fields and unmatched records for manual review

## Workflow

```text
Form-3 Word documents
          +
Timesheet Excel files
          |
          v
Document parsing and field normalization
          |
          v
Student record matching
          |
          v
Payroll field calculations and validation warnings
          |
          v
Institutional Excel templates are populated
          |
          v
Editable Excel outputs for final human review
```

## Technology Stack

* **Python** — application and business logic
* **python-docx** — Word document parsing
* **openpyxl** — Excel reading, writing, and template preservation
* **Flask** — planned web upload and review interface

The current development environment uses Python 3.14.

## Project Structure

```text
Excel-Automation/
├── calc.py              # Timesheet processing, matching, and calculations
├── parser.py            # Form-3 Word document parser
├── writer.py            # Institutional Excel output generation
├── templates.py         # Input-field definitions, template paths, and constants
├── templates/           # Excel workbook templates (local only, not tracked in Git)
├── requirements.txt     # Python dependencies
├── .gitignore
└── README.md
```

## Installation on Windows 11

### 1. Clone the repository

```powershell
git clone https://github.com/neohitaf/Excel-Automation.git
cd Excel-Automation
```

### 2. Create a virtual environment

```powershell
py -m venv .venv
```

### 3. Activate the environment

```powershell
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks the activation script, run the following command once and then activate the environment again:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 4. Install dependencies

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Input, Output, and Template Directories

Create the local working directories below. They are intentionally excluded from Git because they may contain personal or institutional data.

```text
Excel-Automation/
├── veri/
│   ├── Form-3.docx
│   └── Puantaj_ornek.xlsx
├── cikti/
└── templates/
    ├── bordro_sablon.xlsx
    ├── banka_listesi_sablon.xlsx
    └── muhtasar_sablon.xlsx
```

The current command-line workflow expects the input paths defined in `calc.py` and writes generated workbooks to the `cikti/` directory. Template workbooks referenced by `templates.py` must be placed locally in `templates/` and are never committed to version control (see **Privacy and Data Protection** below).

## Running the Current Version

After activating the virtual environment:

```powershell
python calc.py
```

Before running the program, confirm that:

* The expected Form-3 and timesheet files are available in the local `veri/` directory
* The required workbook templates are available locally in `templates/`
* The output directory exists
* Template names and worksheet names have not been changed

## Data Validation Philosophy

The application does not silently redesign or replace the administrative process. It automates repeatable steps and surfaces records that may require human attention.

Examples include:

* Missing mandatory fields
* Form-3 records without a matching timesheet record
* Timesheet records without a matching Form-3 record
* Possible name-format differences
* Unexpected GSS values
* Template-capacity or structural problems

The final Excel files remain editable so administrative staff can perform the established final verification and correction steps.

## Privacy and Data Protection

This workflow processes personally identifiable information such as names, identity numbers, and IBANs.

For that reason:

* Real input documents must not be committed to Git
* Generated output files must not be committed to Git
* **Institutional Excel templates must not be committed to Git**, even with sample data
* Public template files, if ever shared, must contain only empty or synthetic sample data
* Logs and screenshots must not expose personal information
* Local test data should be anonymized before sharing the repository

The `veri/`, `cikti/`, and `templates/` directories are excluded through `.gitignore` for this purpose.

## Current Status

**Active development**

The core command-line workflow for document parsing, record matching, calculations, and Excel output generation is implemented. Validation controls and the user interface are being improved incrementally.

This repository represents an independently developed administrative workflow automation project rather than a finished commercial payroll product.

## Planned Web Interface

The next stage is a lightweight Flask-based interface that will allow users to:

1. Upload Form-3 and timesheet files
2. Select or confirm the payroll period
3. Review matched and unmatched student records
4. See warnings before output generation
5. Generate the institution-compatible workbooks
6. Download the results individually or as a ZIP archive

The web interface will act as an upload, validation, and output-generation layer. Final manual corrections will continue to be made in Excel so the existing institutional process remains intact.

## Planned Improvements

* Stronger validation for GSS, identity-number, and IBAN fields
* Consistent Turkish-character normalization during record matching
* Template-capacity and worksheet-structure checks
* Clearing old input cells without modifying formula cells
* Automated tests using fully synthetic data
* Configurable payroll periods and institution-specific parameters
* Flask-based upload and review workflow
* Downloadable warning and processing reports

## Skills Demonstrated

This project combines software development with business-process analysis and demonstrates:

* Requirements analysis in an existing administrative environment
* Balancing automation with user control and institutional constraints
* Word and Excel document integration
* Data extraction, normalization, and record matching
* Rule-based payroll processing
* Template-compatible output generation
* Error reporting and validation design
* Incremental modernization of a manual workflow

## Disclaimer

This project is designed for workflow automation and educational development. Generated files must be reviewed by authorized personnel before being used in official payroll, tax, banking, or social-security procedures.
