# Solve Analytics Pipeline

A small Python utility that imports **csTimer solve exports into an existing Excel workbook** and automatically calculates rolling performance metrics.

The workflow is simple:

> **Practice/Solve → Export from csTimer → Run the script → Excel is updated**

## Features

* Automatically finds the latest `cstimer*.txt` export.
* Imports new solves into an existing Excel workbook.
* Calculates:

  * Ao5
  * Ao12
  * Ao100
  * 100-solve consistency (population standard deviation)
* Preserves PLL information from the csTimer export.
* Detects which solves are already in the workbook and only adds new ones.
* Fills missing historical metrics without overwriting existing values.
* Adds a summary block for each new batch of solves.
* Supports a safe `--dry-run` mode.

---

## Requirements

* Python 3.9+
* A csTimer export
* An existing Excel workbook with the expected layout

Install the Python dependencies:

```bash
pip install pandas openpyxl
```

Using a virtual environment is recommended:

```bash
python -m venv venv
venv\Scripts\activate
pip install pandas openpyxl
```

`venv/` should be included in `.gitignore`.

---

## Configuration

Create a `config.json` file next to `convert.py`:

```json
{
    "downloads_dir": "PATH_TO_DOWNLOADS",
    "excel_path": "PATH_TO_EXCEL_WORKBOOK",
    "sheet_name": "EXCEL_SHEET_NAME"
}
```

The configuration file contains machine-specific paths, so it is recommended to keep it out of version control:

```gitignore
config.json
venv/
__pycache__/
```

You can optionally provide a `config.example.json` with placeholder paths so other users know what configuration is required.

---

## Usage

### 1. Export your solves

Export your solves from [csTimer](https://cstimer.net/) and place the resulting `.txt` file in the configured downloads directory.

The script looks for files matching:

```text
cstimer*.txt
```

and automatically uses the most recently modified export.

### 2. Run the script

```bash
python convert.py
```

The script will:

1. Find the latest csTimer export.
2. Parse the solve data.
3. Compare it with the existing Excel workbook.
4. Append any new solves.
5. Calculate rolling metrics.
6. Update the summary section.
7. Save the workbook.

### Preview changes without modifying Excel

```bash
python convert.py --dry-run
```

This is useful for verifying that the correct export is being detected and that new solves are being found.

---

## Command-Line Overrides

Configuration values can also be supplied directly:

```bash
python convert.py --downloads-dir "PATH_TO_DOWNLOADS"
```

```bash
python convert.py --excel-path "PATH_TO_WORKBOOK"
```

```bash
python convert.py --sheet-name "SHEET_NAME"
```

Command-line arguments take priority over values in `config.json`.

---

## How the Metrics Work

The script uses rolling solve history to calculate:

| Metric          | Description                                                        |
| --------------- | ------------------------------------------------------------------ |
| **Ao5**         | Average of the most recent 5 solves, dropping the best and worst solves |
| **Ao12**        | Average of the most recent 12 solves, dropping the best and worst solves |
| **Ao100**       | Average of the most recent 100 solves, dropping the best 5 and worst 5 solves |
| **StdDev(100)** | Population standard deviation of the most recent 100 solves        |

Metrics are calculated from the full solve history, so importing a new batch does not reset the rolling averages.

---

## Excel Layout

The script expects the solve data to use the following columns:

```text
A  Solve Number
B  Time
C  PLL
D  Ao5
E  Ao12
F  Ao100
G  100-solve StdDev
```

The workbook is therefore expected to follow the structure of the author's existing template rather than being a completely generic Excel file.

---

## Assumptions & Limitations

* The csTimer export contains the expected session data (`session1` by default).
* Solve times are available as numeric values.
* The workbook already exists and contains the expected worksheet/layout.
* DNFs are not currently modeled separately.
* Solve numbers are used to determine which solves are new.

This project is primarily intended as a **personal automation tool for maintaining a long-term speedcubing practice log**, rather than a general-purpose csTimer analytics application.

---

## Project Structure

```text
.
├── convert.py
├── README.md
├── config.json          # local only, not committed
├── requirements.txt
└── venv/                # local only, not committed
```

For sharing the project, consider including a `config.example.json` instead of the real `config.json`.

---

## Future Ideas

Possible extensions that I may add in the future:

* DNF-aware WCA averages
* Additional statistics and charts
* PLL frequency analysis
* Progress tracking over time
* Better duplicate detection
* Automated Excel backups
* Support for multiple csTimer sessions