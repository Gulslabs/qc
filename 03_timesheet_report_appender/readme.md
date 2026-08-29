# Attendance Time Sheet Report AppenderBin Analyzer

A Python tool that appends student attendance data by creating time-based bins and calculating frequency distributions for join and leave times. 

### Required Python Packages

```bash
pip install pandas openpyxl
```

**Example Data:**
```
Name                              Join Time    Leave Time
AJ_007_Khalid_Azeem              20:36        22:48
AJ_014_ABDULRAHMAN               20:41        22:47
AJ_013_vizarath_Ali              20:42        22:47
```

### Usage Command Line (Recommended)

```bash
python timesheet_report_appender_v1.py Time_Sheet_KARMH-B02_Surah_Al-Ambiya.xlsx 3
python timesheet_report_appender_v1.py Time_Sheet_KARMH-B02_Surah_Az-Zukhruf_22AUG2026.xlsx 5
python  attendance_timeline.py Time_Sheet_KARMH-B02_Surah_Ha-Mim_Sajda_01AUG2026.xlsx 5
```


### Generated Excel File

The tool creates a new Excel file named `<original_name>_with_reports.xlsx` with 6 additional columns:

| Column | Description |
|--------|-------------|
| `Bin_Start_Join_Time` | Start time of join time bin |
| `Bin_End_Join_Time` | End time of join time bin |
| `Join_Student_Count` | Number of students who joined in this time range |
| `Bin_Start_Leave_Time` | Start time of leave time bin |
| `Bin_End_Leave_Time` | End time of leave time bin |
| `Leave_Student_Count` | Number of students who left in this time range |

**Example Output:**
```
Bin_Start_Join_Time  Bin_End_Join_Time  Join_Student_Count
20:30                20:35              11
20:35                20:40              15
20:40                20:45              19
20:45                20:50              49
```
### Graph (Manual)
X-Axis Value: `Time Range (3-min intervals)`
Y-Axis Value: `Student Count`