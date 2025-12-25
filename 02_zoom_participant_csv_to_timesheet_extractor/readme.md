# Zoom CSV to Timesheet Converter

A Python tool that converts Zoom participant CSV files into formatted Excel timesheets with attendance tracking and analysis features.

### Requirements
Create a `requirements.txt` file:
```
pandas==2.1.3
openpyxl==3.1.2
```
```bash
pip install -r requirements.txt
```

### Basic Command
```bash
python.exe .\csv_to_timesheet_extractor_v3.py .\participants_81868085400.csv .\naqeeb_to_initial_KARMH-02.csv "20:50" "22:50"
python.exe .\csv_to_timesheet_extractor_v3.py .\participants_2816472229.csv .\naqeeb_to_initial_TSAP-02.csv "20:50" "22:40"
```

### Parameters
- **CSV File**: Zoom participant export file
- **naqeeb_mapping.csv**: A csv file with naqeeb name, his initial and his Student Id Range
- **Start Time**: Session start time in 24-hour format (HH:MM)
- **End Time**: Session end time in 24-hour format (HH:MM)



## Features

### 1. Name Processing
Automatically formats participant names based on role:
- `Tariq Fareed` → `Tariq Fareed` (unchanged)
- `IK_Irfan Khan(Naqeeb) (Full Name)` → `IK_Irfan Khan(Naqeeb)` (keeps Naqeeb title)
- `U1615 AZHARULLAH (Other Info)` → `U1615 AZHARULLAH` (removes non-Naqeeb brackets)

### 2. Duplicate Consolidation
Merges multiple join/leave sessions for same participant:
- **Before**: 
  ```
  Inam Khan(Naqeeb)  07:56:17 PM  09:16:35 PM   81 minutes
  Inam Khan(Naqeeb)  09:16:36 PM  10:31:22 PM   75 minutes
  Inam Khan(Naqeeb)  10:31:22 PM  10:37:00 PM    6 minutes
  Inam Khan(Naqeeb)  10:37:00 PM  11:01:37 PM   25 minutes
  Inam Khan(Naqeeb)  11:01:38 PM  11:13:16 PM   12 minutes
  ```
- **After**: 1 record with join time `07:56:17 PM`, leave time `11:13:16 PM`, total duration `199 minutes` (81+75+6+25+12)

### 3. Naqeeb Name Mapping
Maps participant initials to Naqeeb names and ID range; using `naqeeb_to_initial.csv`:
```csv
Ayyaz Sayyed,AZ
Rais Tigadi,RT
```
- `AZ_Participant Name` → Naqeeb Name: "Ayyaz Sayyed"
- `RT K599 Someone` → Naqeeb Name: "Rais Tigadi"

### 4. Attendance Tracking
Marks attendance issues based on configurable thresholds (default: 5 minutes):
- **Late Joiner**: Joined 5+ minutes after start time
- **Early Leaver**: Left 5+ minutes before end time
- **Both**: Amber highlighting for single issues, red for both

### 5. Time Format Conversion
Converts Zoom's datetime format to clean 24-hour time:
- `8/23/2025 07:35:40 PM` → `19:35`
- `8/23/2025 11:13:17 PM` → `23:13`

### 7. Conditional Highlighting
  - Amber background: Single attendance issue
  - Red background: Multiple attendance issues

### 8. Configuration
```python
ENABLE_DURATION_THRESHOLD = False  # Enable/disable sheet splitting
DURATION_THRESHOLD = 20           # Minimum minutes for main sheet
START_TIME_THRESHOLD_MINUTES = 5  # Late joiner threshold
END_TIME_THRESHOLD_MINUTES = 5    # Early leaver threshold
```

### 9. Sample Output
```
Name                    Join Time  Leave Time  Duration  Naqeeb Name    Remarks
SR_Samir Shaikh(Naqeeb) 19:35      22:32       177      Samir Shaikh   
RT_Participant Name     20:50      22:55       125      Rais Tigadi    Late Joiner
AZ_Another Person       19:30      22:50       200      Ayyaz Sayyed   Early Leaver
```
