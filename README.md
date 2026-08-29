# Project under this repository

This repository contains scripts for attendance extraction, Zoom participant conversion, and timesheet reporting.

## Download this project from GitHub

1. Open `https://github.com/Gulslabs/qc` GitHub repository page.
2. Click the green "Code" button.
3. Select "Download ZIP".
4. Save the ZIP file to your computer, for example to your Desktop or Downloads folder.
5. After download, right-click the ZIP file and choose "Extract All".
6. Choose a folder such as `C:\Users\YourName\Downloads\qc` or a folder on your Desktop.
7. After extraction, open that folder in File Explorer.

You should now see files like:
- `README.md`
- `requirements.txt`
- `01_attendance_extractor/`
- `02_zoom_participant_csv_to_timesheet_extractor/`
- `03_timesheet_report_appender/`

## Install Python on Windows

1. Download Python from the official site: https://www.python.org/downloads/windows/
2. Run the installer.
3. Check the box: "Add Python to PATH".
4. If you see "Install launcher for all users", keep it checked.
5. Finish the installation.
6. Open a new Command Prompt or PowerShell window and verify Python is installed:

```bash
python --version
```

If that does not work, try:

```bash
py --version
```

## Install required packages

This repo includes a single requirements file so setup is simple:

```bash
py -m pip install -r requirements.txt
```

This installs the packages needed by the scripts in the project folders.

## 1) Attendance Extractor
Location: `01_attendance_extractor`

Run the program:

```bash
python.exe attendance_extractor_2.3.py "D:\Work\py-ws\qc\01_attendance_extractor\TSAP-B02" "24-12-25" "TSAP-B02"
```

This script reads attendance files from a folder and creates a consolidated master sheet.

Input parameters:
- `parent_folder`: Folder containing the Naqeeb attendance sheets.
- `attendance_date`: Date to extract, such as `22-08-26`, `07-01-26`, or `23 Dec 25`.
- `id_prefix`: ID filter used to keep matching records, such as `KARMH-B02`, `TSAP-B02`, or `Attendance`.

Example:

```bash
python.exe attendance_extractor_2.2.py "D:\Work\py-ws\qc\01_attendance_extractor\KARMH-B02" "22-08-26" "KARMH-B02"
```

---

## 2) Zoom Participant CSV to Timesheet Extractor
Location: `02_zoom_participant_csv_to_timesheet_extractor`

Run the program:

```bash
python.exe .\csv_to_timesheet_extractor_v3.py .\participants_89078282589.csv .\naqeeb_to_initial_KARMH-02.csv "21:05" "22:50"
```

This script converts a Zoom participant export CSV into a formatted Excel timesheet with attendance checks and mapping to Naqeeb names.

Input parameters:
- `CSV File`: Zoom participant export file.
- `naqeeb_mapping.csv`: CSV file with Naqeeb names, initials, and student ID range.
- `Start Time`: Session start time in 24-hour format, for example `21:05`.
- `End Time`: Session end time in 24-hour format, for example `22:50`.

---

## 3) Timesheet Report Appender
Location: `03_timesheet_report_appender`

Run the program:

```bash
python timesheet_report_appender_v1.py Time_Sheet_KARMH-B02_Surah_Az-Zukhruf_22AUG2026.xlsx 5
```

This script appends time-based attendance summary reports to an Excel timesheet by grouping join and leave times into bins.

Input parameters:
- `Excel File`: The timesheet Excel file to process.
- `bin_size_minutes`: The size of each time bin in minutes, for example `5`.

---

## Quick reference

- Attendance extraction: `python.exe attendance_extractor_2.3.py "<folder>" "<date>" "<id_prefix>"`
- Zoom CSV conversion: `python.exe .\csv_to_timesheet_extractor_v3.py .\<participants.csv> .\<naqeeb_mapping.csv> "<start_time>" "<end_time>"`
- Report appender: `python timesheet_report_appender_v1.py <excel_file> <bin_size_minutes>`
