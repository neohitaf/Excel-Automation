# Student Payroll Excel Automation

A Python-based administrative workflow automation project for processing part-time student employment records and generating institution-compatible Excel outputs.

The project was developed to reduce repetitive manual data entry in a university administrative workflow while preserving the institution's existing Excel templates, formulas, formatting, and manual review process.

## Project Context

The workflow uses two main source-document types with different roles:

- **Form-3 Word document** — a relatively stable master-data source containing student identity, banking, IBAN, GSS, and related personnel information
- **Monthly timesheet Excel files** — recurring documents containing the unit, role, payroll period, and total working-hour information required for monthly payroll processing

The monthly timesheet is matched with the Form-3 records to generate payroll-related outputs such as the payroll workbook, tax declaration workbook, and bank list.

The institution also uses approved Excel templates with predefined layouts, formulas, sections, and manual control points. Recreating these workbooks dynamically would risk changing the established administrative process. For this reason, the project writes data into the approved templates rather than replacing them.

## Design Approach

This project intentionally follows a **template-preserving and human-in-the-loop** approach.

The output workbooks are not recreated from scratch. Instead, the application opens the institution's existing Excel templates and writes data only to designated input cells.

This approach preserves:

- Existing workbook layouts and formatting
- Institution-defined formulas
- Formula cells that must remain visible even when their result is `0`
- Fixed sections used for different employee document types
- Existing manual review and correction procedures
- The ability of authorized staff to adjust the final workbooks when necessary

The fixed row positions and workbook-specific rules in the code represent the structure of approved institutional templates rather than a generic dynamic spreadsheet design.

## Current Features

- Parses Form-3 `.docx` files and locates the relevant student table
- Supports alternative column-title variations in source documents
- Extracts student identity, GSS, bank, IBAN, and related information
- Reads monthly `.xlsx` timesheet files and identifies student records
- Extracts unit names, payroll periods, and total working hours
- Normalizes and matches Form-3 records with timesheet records
- Uses different source priorities according to output requirements:
  - The **bank list** is primarily Form-3-based and is enriched with unit information from the timesheet
  - The **payroll** and **tax declaration** outputs are primarily timesheet-based and are enriched with Form-3 information
- Calculates premium days, missing days, and payroll-related fields
- Separates records according to GSS status and document type
- Preserves formulas, formatting, and editable cells in institutional templates
- Reports missing fields and unmatched records for manual review
- Generates the following Excel outputs:
  - `bordro.xlsx`
  - `muhtasar.xlsx`
  - `banka_listesi.xlsx`

## Workflow

```text
Stable Form-3 master document
            +
Monthly timesheet Excel file
            |
            v
Document parsing and field normalization
            |
            v
Student record matching and data enrichment
            |
            v
Payroll calculations and validation warnings
            |
            v
Institutional Excel templates are populated
            |
            v
Editable Excel outputs for final human review
```

## Application Architecture

The project separates the automation logic from the web layer.

```text
FastAPI
  |
  |-- receives uploaded files
  |-- validates HTTP requests
  |-- calls the automation service
  |-- returns result and download responses
  |
  v
calistir()
  |
  |-- parses Form-3
  |-- reads the monthly timesheet
  |-- matches and enriches records
  |-- calculates payroll fields
  |-- generates Excel outputs
  |-- returns warnings and output paths
  |
  v
parser.py + calc.py + writer.py
```

The core `calistir()` function is independent of FastAPI. It accepts input and output paths and returns a structured result:

```python
{
    "uyarilar": [...],
    "cikti_dosyalari": [
        "banka_listesi.xlsx",
        "bordro.xlsx",
        "muhtasar.xlsx",
    ],
}
```

This allows the same automation workflow to be used from the command line, a web interface, or future integrations without rewriting the business logic.

## Web Interface Stack

The planned local web interface uses:

- **FastAPI** — HTTP layer, file upload and download routes, and automation-service calls
- **Jinja2** — server-side HTML templates for upload, result, and error pages
- **Bootstrap** — responsive forms, alerts, tables, buttons, and page layout
- **Vanilla JavaScript** — small user-experience enhancements such as displaying selected filenames and disabling the submit button during processing

The web interface is intended to remain lightweight. The Excel workbooks continue to be the final review and correction environment.

## Technology Stack

- **Python** — application and business logic
- **python-docx** — Word document parsing
- **openpyxl** — Excel reading, writing, and template preservation
- **pandas** — structured data-processing support where required
- **FastAPI** — local web application and HTTP endpoints
- **Jinja2** — server-side HTML rendering
- **Bootstrap** — responsive interface components
- **Vanilla JavaScript** — basic client-side interaction

The current development environment uses Python 3.14.

## Project Structure

The target structure for the application is:

```text
Excel-Automation/
├── app/
│   ├── main.py                 # FastAPI application and HTTP routes
│   ├── service.py              # Main calistir() automation workflow
│   ├── parser.py               # Form-3 Word document parser
│   ├── calc.py                 # Timesheet reading, matching, and calculations
│   ├── writer.py               # Institutional Excel output generation
│   ├── templates.py            # Field definitions, workbook paths, and constants
│   ├── web_templates/
│   │   ├── index.html          # File-upload page
│   │   ├── result.html         # Processing result and warning page
│   │   └── error.html          # User-facing error page
│   └── static/
│       └── js/
│           └── app.js          # Minimal client-side interactions
├── institutional_templates/   # Local Excel templates, not tracked in Git
├── veri/                       # Local input documents, not tracked in Git
├── cikti/                      # Generated workbooks, not tracked in Git
├── instance/
│   └── jobs/                   # Temporary per-request working directories
├── tests/
├── requirements.txt
├── .gitignore
└── README.md
```

The exact folder names may change during development, but the separation between the web layer, automation service, document-processing modules, and institutional templates will remain.

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

## Local Input, Output, and Template Directories

Create the local working directories below. They are intentionally excluded from Git because they may contain personal or institutional data.

```text
Excel-Automation/
├── veri/
│   ├── Form-3.docx
│   └── Puantaj_ornek.xlsx
├── cikti/
└── institutional_templates/
    ├── bordro_sablon.xlsx
    ├── banka_listesi_sablon.xlsx
    └── muhtasar_sablon.xlsx
```

The Form-3 document acts as a master-data source and may remain unchanged across multiple monthly payroll periods. The timesheet is a recurring monthly input and is processed together with the active Form-3 data.

Template workbooks must be stored locally and must never be committed to version control.

## Running the Command-Line Workflow

The command-line entry point calls the reusable `calistir()` function with local paths.

Example:

```python
from service import calistir

sonuc = calistir(
    "veri/Form-3.docx",
    "veri/Puantaj_ornek.xlsx",
    "cikti",
)
```

A typical result contains warnings and generated output paths:

```python
{
    "uyarilar": [
        "A timesheet record could not be matched with Form-3."
    ],
    "cikti_dosyalari": [
        "cikti/banka_listesi.xlsx",
        "cikti/bordro.xlsx",
        "cikti/muhtasar.xlsx",
    ],
}
```

The `if __name__ == "__main__":` block is used only as a local command-line entry point. The web layer will call the same `calistir()` function with request-specific file paths.

## Planned Local Web Workflow

The FastAPI-based interface will provide the following flow:

1. Select the Form-3 document
2. Select the monthly timesheet
3. Upload and process the files
4. Display the detected payroll period
5. Show matched, unmatched, and warning records
6. Generate institution-compatible workbooks
7. Download the outputs individually or as a ZIP archive

The first web release may require both files to be selected for each operation. A later version may support storing and versioning an active Form-3 document so that users only upload the monthly timesheet.

## Data Validation Philosophy

The application does not silently redesign or replace the administrative process. It automates repeatable steps and surfaces records that may require human attention.

Warnings represent conditions where processing may continue but the final workbook should be checked.

Examples include:

- Missing optional or reviewable fields
- Form-3 records without a matching monthly timesheet record
- Timesheet records without a matching Form-3 record
- Possible name-format or Turkish-character differences
- Unexpected GSS values
- Empty unit information
- Records requiring manual confirmation

Critical validation errors should stop output generation.

Examples include:

- An unreadable Form-3 or timesheet file
- Missing mandatory columns
- An undetectable payroll period
- Missing institutional templates
- No processable records
- Workbook structures that do not match expected templates
- Output files that cannot be written safely

The final Excel files remain editable so administrative staff can perform the established verification and correction steps.

## Privacy and Data Protection

This workflow processes personally identifiable information such as names, identity numbers, and IBANs.

For that reason:

- Real input documents must not be committed to Git
- Generated output files must not be committed to Git
- Institutional Excel templates must not be committed to Git, even when they contain sample data
- Public sample templates must contain only empty or fully synthetic data
- Logs, screenshots, and demonstrations must not expose personal information
- Local test data must be anonymized before repository publication
- Uploaded web files should be stored only in temporary, request-specific directories
- Temporary files should be deleted after the processing or download workflow is complete
- The first deployment should remain local unless the institution provides an approved secure hosting environment

The `veri/`, `cikti/`, `institutional_templates/`, and temporary job directories should be excluded through `.gitignore`.

## Current Status

**Active development**

The core workflow for document parsing, timesheet reading, record matching, payroll calculations, and Excel output generation is implemented.

The current development focus is:

- Finalizing the reusable `calistir()` service flow
- Improving warning and validation results
- Creating the FastAPI upload and download layer
- Rendering processing results through Jinja2 templates
- Building a lightweight Bootstrap interface
- Adding small client-side interactions with vanilla JavaScript
- Testing the workflow with anonymized and synthetic data

This repository represents an independently developed administrative workflow automation project rather than a finished commercial payroll product.

## Planned Improvements

- FastAPI-based upload, processing, and download workflow
- Jinja2 result and error pages
- Bootstrap-based responsive interface
- Downloadable ZIP packages and processing reports
- Stronger validation for GSS, identity-number, and IBAN fields
- Consistent Turkish-character normalization during record matching
- Template-capacity and worksheet-structure checks
- Clearing old input cells without modifying formula cells
- Automated tests using fully synthetic data
- Configurable payroll periods and institution-specific parameters
- Active Form-3 storage and version tracking
- Per-request temporary-directory cleanup
- Improved processing statistics for matched and unmatched records

## Skills Demonstrated

This project combines software development with business-process analysis and demonstrates:

- Requirements analysis in an existing administrative environment
- Balancing automation with user control and institutional constraints
- Object-oriented and modular software design principles
- Word and Excel document integration
- Data extraction, normalization, and record matching
- Rule-based payroll processing
- Template-compatible output generation
- Structured warning and result handling
- FastAPI-based backend development
- Server-side HTML rendering with Jinja2
- Responsive interface development with Bootstrap
- Basic client-side interaction with vanilla JavaScript
- Incremental modernization of a manual institutional workflow

## Disclaimer

This project is designed for workflow automation and educational development.

Generated files must be reviewed by authorized personnel before being used in official payroll, tax, banking, or social-security procedures. The application does not replace institutional approval, accounting controls, or legal verification processes.
