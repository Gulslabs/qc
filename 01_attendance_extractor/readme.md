# Attendance Master Sheet Generator

A Python tool to extract and consolidate attendance data from multiple Naqeeb attendance sheets into a single master sheet, with filtering by ID prefix.

## Features

- ✅ Filters records by ID prefix (e.g., `KARMH`, `KARMH-B02`)
- ✅ Automatically moves processed files to archive folder
- ✅ Performance optimized for large datasets

## Usage Basic Command

```bash
python attendance_generator.py <parent_folder> <attendance_date> <id_prefix>
```
**Extract all KARMH-B02 records for December 20, 2025:**
```bash
python attendance_generator.py "C:\Users\admin\Documents\Attendance" 20-12-25 KARMH
```
### Parameters
| Parameter | Description | Example |
|-----------|-------------|---------|
| `parent_folder` | Path to folder containing Naqeeb attendance sheets | `/path/to/attendance/files` |
| `attendance_date` | Date column to extract (format: DD-MM-YY or DD-MM-YYYY) | `20-12-25` |
| `id_prefix` | Filter records by ID prefix | `KARMH` or `KARMH-B02` |

## Output Generated Master Sheet

The tool creates a file named `master_sheet_<date>.xlsx` with the following structure:

| ID# Assigned | Region | Naqeebs | 20-12-25 |
|--------------|--------|---------|----------|
| KARMH-B02-G1401 | GUJ | Parvez Misarwala | P |
| KARMH-B02-G1402 | GUJ | Parvez Misarwala | A |
| KARMH-B02-M1223 | MH | Ayyaz Sayyed | P |

### Attendance Values

- `P` - Present
- `A` - Absent
- `L` - Leave
- `D` - Dropout
- `-` - Not marked / Default

### Processed Files

Original source files are moved to `Processed/` folder with timestamp:
```
Original: Abdul Haq Naqeeb Sheet (KARMH-B02).xlsx
Moved to: Processed/Abdul Haq Naqeeb Sheet (KARMH-B02)_14_30.xlsx
```

## Folder Structure

```
parent_folder/
├── Naqeeb Sheet 1.xlsx
├── Naqeeb Sheet 2.xlsx
├── Naqeeb Sheet 3.xlsx
├── master_sheet_20-12-25.xlsx    (generated output)
└── Processed/
    ├── Naqeeb Sheet 1_14_30.xlsx
    ├── Naqeeb Sheet 2_14_30.xlsx
    └── Naqeeb Sheet 3_14_30.xlsx
```

## Performance Optimizations

- **Read-only mode**: Faster file loading, lower memory usage
- **Write-only mode**: Faster output generation
- **Single-pass header scanning**: Reduced column lookups
- **Batch row processing**: Efficient data writing
- **Value-only iteration**: No unnecessary cell object creation

### Dependencies

- `openpyxl>=3.1.2` - Excel file handling
- `pathlib` - Cross-platform path operations
- `datetime` - Timestamp generation
- `shutil` - File operations

*** How to run: 
`python.exe attendance_extractor_2.2.py "D:\Work\py-ws\qc\01_attendance_extractor\KARMH-B02" "07-03-26" "KARMH-B02"`
`python.exe attendance_extractor_2.2.py "D:\Work\py-ws\qc\01_attendance_extractor\TSAP-B02" "07-01-26" "TSAP-B02"` # Without Comments
`python.exe attendance_extractor_2.2.py "D:\Work\py-ws\qc\01_attendance_extractor\HYDLH-B01" "23 Dec 25" "Attendance"` # Without Comments
`python.exe attendance_extractor_2.2.py "D:\Work\py-ws\qc\01_attendance_extractor\HYDLH-B01" "23`nDec`n25" "Attendance"` # Without Comments
`python.exe attendance_extractor_2.3.py "D:\Work\py-ws\qc\01_attendance_extractor\TSAP-B02" "24-12-25" "TSAP-B02"` # With Comments

"23
Dec
25"
"23
Dec
25"
