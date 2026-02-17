# Document Ageing Report Automation

This project automates the generation of document ageing reports from Excel export files.

## Prerequisites

- Python 3.8 or higher
- pip (Python package installer)

## Setup

### 1. Create a Virtual Environment

**Windows:**
```powershell
python -m venv venv
```

**Linux/Mac:**
```bash
python3 -m venv venv
```

### 2. Activate the Virtual Environment

**Windows (PowerShell):**
```powershell
.\venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**
```cmd
venv\Scripts\activate.bat
```

**Linux/Mac:**
```bash
source venv/bin/activate
```

### 3. Install Dependencies

Once the virtual environment is activated, install the required packages:

```bash
pip install -r requirements.txt
```

## Running the Application

### Basic Usage

Run the application with default settings:

```bash
python main.py
```

This will:
- Read from `data/export.xls` (default input file)
- Generate the report in the `output` directory
- Save logs to the `logs` directory

### Command-Line Options

You can customize the behavior using command-line arguments:

```bash
python main.py --input <path_to_input_file> --output-dir <output_directory> --logs-dir <logs_directory> --log-level <LOG_LEVEL>
```

**Available Options:**
- `--input`: Path to the input Excel file (default: `data/export.xls`)
- `--output-dir`: Directory to save the generated workbook (default: `output`)
- `--logs-dir`: Directory to save log files (default: `logs`)
- `--log-level`: Logging level - DEBUG, INFO, WARNING, or ERROR (default: `INFO`)

**Examples:**

```bash
# Use a custom input file
python main.py --input data/my_export.xls

# Specify custom output and logs directories
python main.py --output-dir reports --logs-dir app_logs

# Run with debug logging
python main.py --log-level DEBUG

# Combine multiple options
python main.py --input data/custom.xls --output-dir reports --log-level DEBUG
```

## Deactivating the Virtual Environment

When you're done working on the project, you can deactivate the virtual environment:

```bash
deactivate
```

## Project Structure

```
.
├── data/              # Input Excel files
├── output/            # Generated reports
├── logs/              # Log files
├── src/               # Source code modules
│   ├── config.py
│   ├── data_loader.py
│   ├── excel_styles.py
│   ├── logging_setup.py
│   ├── report_builder.py
│   └── runner.py
├── main.py            # Main entry point
├── requirements.txt   # Python dependencies
└── README.md          # This file
```

## Troubleshooting

- **If you get a "permission denied" error when activating venv on Windows PowerShell:**
  - Run PowerShell as Administrator, or
  - Execute: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

- **If dependencies fail to install:**
  - Make sure your virtual environment is activated
  - Ensure you have an active internet connection
  - Try upgrading pip: `python -m pip install --upgrade pip`

- **If the input file is not found:**
  - Check that the file path is correct
  - Use the `--input` argument to specify the correct path
