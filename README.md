# Staj Otomasyon

A Python-based automation tool for processing part-time student employment records at Necmettin Erbakan University.

## What It Does
- Parses Form-3 Word documents to extract student records
- Reads timesheet Excel files to calculate working hours
- Generates payroll outputs: BORDRO, MUHTASAR, BANKA LİSTESİ and more

## Tech Stack
- Python 3.14
- python-docx — Word document parsing
- openpyxl — Excel read/write
- pandas — data processing
- Flask — web interface (in progress)

## Status
🚧 Under active development

## Setup
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Project Structure

cat > README.md << 'EOF'
# Staj Otomasyon

A Python-based automation tool for processing part-time student employment records at Necmettin Erbakan University.

## What It Does
- Parses Form-3 Word documents to extract student records
- Reads timesheet Excel files to calculate working hours
- Generates payroll outputs: BORDRO, MUHTASAR, BANKA LİSTESİ and more

## Tech Stack
- Python 3.14
- python-docx — Word document parsing
- openpyxl — Excel read/write
- pandas — data processing
- Flask — web interface (in progress)

## Status  Under active development

## Setup
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Project Structure
cat > README.md << 'EOF'
# Staj Otomasyon

A Python-based automation tool for processing part-time student employment records at Necmettin Erbakan University.

## What It Does
- Parses Form-3 Word documents to extract student records
- Reads timesheet Excel files to calculate working hours
- Generates payroll outputs: BORDRO, MUHTASAR, BANKA LİSTESİ and more

## Tech Stack
- Python 3.14
- python-docx — Word document parsing
- openpyxl — Excel read/write
- pandas — data processing
- Flask — web interface (in progress)

## Status  Under active development

## Setup
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Project Structure staj-otomasyon/

├── templates.py   # table schema definitions

├── parser.py      # Form-3 document parser

├── calc.py        # payroll calculation engine

├── writer.py      # Excel output generator (in progress)

└── requirements.txt EOF
