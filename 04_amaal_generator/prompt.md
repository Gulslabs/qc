I need a python program for the following. 

Source File: is an Excel that contains; 3 Sheets.  
Sheet 1.  
    Name: 'Askaar'
    Columns:'Timestamp,Generated Naqeeb Name,Week Ending Date,Ba-Wuzu,Tahajjud,Morning Azkaar,Evening Azkaar,Roza,Sadaqah,Did you listen to previous Tarbiya Class Recording?'
Sheet 2: 
    Name: 'Halqa'
    Columns: 'Naqeeb Name,Region,Date,Total Active,PH,P,A,L,TP,A+L,A+L%,Live Halqa Count,Halqa Abs,Halqa Abs%'
Sheet 3: 
    Name: 'Reference Data'
    Columns: 'Srl No,Naqeeb Name,Gmail Id,Current Role,Category'

Output File: contains 
Sheet 1. 
    Name: 'Combined_Report'
    Columns: 'Naqeeb Name,Role,Category,Ba-Wazu,Tahajjud,Azkaar,Roza,Sadqa,Ruju Percentage (33.33%),Book Completion,Listened to Previous Tarbiya Audio,Read Halqa Notes(<48 Hours),Fikr Percentage (33.33%),No of Students,Present,Absent,Leave,Attendance Percentage (33.33%),Overall  Percentage(100%),Week Ending Date'

Objective: 
    An existing Output excel Name 'KAR_Naqeeb_Amaal_And_Seerah_Tracker_Master_Sheet.xlsx'  contains a sheet 'Combined_Report'. The objective is to add 1 row per 'Reference Data'.'Naqeeb Name' for given  `Week_End_Date` into 'Combined_Report'. 

Input Parameters: 
    Param 1: 'Source File'
    Param 2: 'Week_End_Date'
    Param 3: 'Class_Date'


Mapping Logic: 
    Direct columns: 
        Combined_Report.Naqeeb Name = Reference Data.Naqeeb Name
        Combined_Report.Role = Reference Data.Role
        Combined_Report.Ba-Wazu = Askaar.Ba-Wuzu
        Combined_Report.Tahajjud=Askaar.Tahajjud
        Combined_Report.Azkaar = Askaar.Morning Azkaar
        Combined_Report.Roza = Askaar.Roza (If this value is greater than 2; then reset it to 1)
        Combined_Report.Sadqa = Askaar.Sadaqah 
        Combined_Report.Listened to Previous Tarbiya Audio=Askaar.Did you listen to previous Tarbiya Class Recording?
        Combined_Report.Read Halqa Notes(<48 Hours) = Y (Default Value)
        Combined_Report.No of Students = Halqa.Total Active
        Combined_Report.Present = Halqa.TP
        Combined_Report.Absent =  Halqa.A
        Combined_Report.Leave =  Halqa.L
        Combined_Report.Week Ending Date = Input Parameter.Week_End_Date
        Combined_Report.Drop in Halqa = Halqa.Halqa Abs        
    Formulat Columns: 
        Combined_Report.'Ruju Percentage (33.33%)' should contain formula similar to  `=SUM(D1636:H1636)/30*0.33`
        Combined_Report.'Fikr Percentage (33.33%)' should contain formula similar to  `=((K1636*1/3) + IF(L1636="Y",1/3,0) + IF(OR(B1636<>"Class-Naqeeb",M1636="Y"),1/3,0))*0.33`
        Combined_Report.'Direct Attendance Percentage' should contain formula similar to `=IFERROR(IF(B6="Class-Naqeeb",P6/O6,0),0)`
        Combined_Report.'Attendance Percentage (33.33%) should contain formula similar to `=IFERROR(IF([@Role]="Class-Naqeeb",[@Present]/[@[No of Students]]*0.33,0),0)`    
        Combined_Report.'Overall  Percentage(100%)' should contain formula similar to `=IFERROR(SUM(J1636+N1636+T1636),0)`
        Combined_Report.'Category' should contain direct formula similar to `=IFERROR(IF([@[Overall  Percentage(100%)]]>=0.91,"L1",IF([@[Overall  Percentage(100%)]]>=0.8,"L2","L3")),"L3")`


Addtional Conditions: 
    1. Search the Input source file 'Askaar' sheet rows for give Reference Data.Naqeeb Name and Input Parameter.Week_End_Date; to find the correct row; to apply mapping logic. 
    2. Search the Input source file 'Halqa' sheet rows for Reference Data.Naqeeb Name and Input Parameter.Class_Date;to find the correct row; to apply mapping logic. THe halqa table only contains records for Naqeeb Name who's role is `Class-Naqeeb` in reference Data
    3. A row per Reference Data.Naqeeb Name and Input Parameter.Week_End_Date must be added even if data is not available in 'Askaar' or 'Halqa' sheets of input file. 
    4. Output File and Input source file are available locally project directory. 
    5. Oput file is huge file it contains existing data from row 5. So this process should append new rows. 


Data:
Input Source File. 'Askaar' Sheet
```
Timestamp   Generated Naqeeb Name   Week Ending Date    Ba-Wuzu Tahajjud    Morning Azkaar  Evening Azkaar  Roza    Sadaqah Did you listen to previous Tarbiya Class Recording? Email Address
7/22/2026 21:39:33  Irfan Bhai  19-Jul-26   7   6   7   7   0   7   N   y.irfan.khan@gmail.com
7/26/2026 12:09:12  Aamir/Muaz Bhai 26-Jul-26   7   4   7   5   1   7   Y   shaikmuaz92@gmail.com
7/26/2026 18:12:22  Abdul Haq Bhai  26-Jul-26   7   5   7   7   1   7   N   nalband.abdulhaq@gmail.com
7/27/2026 17:06:53  Ansaar Bhai 26-Jul-26   7   6   6   6   0   7   Y   rsa5ansar@gmail.com
```

Input Source File.'Halqa' Sheet
```
Naqeeb Name Region  Date    Total Active    PH  P   A   L   TP  A+L A+L%    Live Halqa Count    Halqa Abs   Halqa Abs%
Abdul Haq   KAR 08-08-2026  22  0   21  0   1   21  1   0.045454545 20  1   
Ansar Ahmed KAR 08-08-2026  20  0   14  2   4   14  6   0.3 13  1   
Faheem Abdullah TN  08-08-2026  0   0   0   0   0   0   0   #DIV/0!     0   
Faruk Pasha KAR 08-08-2026  22  0   19  0   3   19  3   0.136363636 17  2   
Karimullah  KAR 08-08-2026  21  0   16  1   4   16  5   0.238095238 13  3   
```
Input Source File.'Reference' Sheet
```
Srl No  Naqeeb Name Gmail Id    Current Role    Category
1   Abdul Haq Bhai  nalband.abdulhaq@gmail.com  Class-Naqeeb    L3
2   Ahsan Bhai  gulam.ahsan@gmail.com   IT/Support  L3
3   Ansaar Bhai rsa5ansar@gmail.com     Class-Naqeeb    L3
4   Fahim Abdullah Bhai Faheem.abdullah@gmail.com   Class-Naqeeb    L3
```

Output File.Combined_Report Sheet

```
Naqeeb Name Role    Category    Ba-Wazu Tahajjud    Azkaar  Roza    Sadqa   Zikr    Ruju Percentage (33.33%)    Book Completion Listened to Previous Tarbiya Audio  Read Halqa Notes(<48 Hours) Fikr Percentage (33.33%)    No of Students  Present Absent  Leave   Drop in Halqa   Attendance Percentage (33.33%)  Direct Attendance Percentage    Overall  Percentage(100%)   Week Ending Date
Abdul Haq Bhai  Class-Naqeeb    L3  7   5   7   0   7       28.60%  40% N   Y   15.40%  22  21  0   1       31.50%  95.5%   75.50%  09-Aug-26
Ahsan Bhai  IT/Support  L3  7   6   6   0   7       28.60%  100%        Y   22.00%                      0.00%   0.0%    50.60%  09-Aug-26
Ansaar Bhai Class-Naqeeb    L3  7   5   6   1   7       28.60%  47% N   Y   16.22%  20  14  2   4       23.10%  70.0%   67.92%  09-Aug-26
Fahim Abdullah Bhai Class-Naqeeb    L2  7   5   5   0   5       24.20%  50% Y   Y   27.45%  22  20  0   2       30.00%  90.9%   81.65%  09-Aug-26
Faruk Pasha Bhai    Class-Naqeeb    L2  7   4   7   1   7       28.60%  55% Y   Y   28.04%  22  19  0   3       28.50%  86.4%   85.14%  09-Aug-26
Irfan Bhai  Class-Naqeeb    L3                          0.00%   42%     Y   15.64%                      0.00%   0.0%    15.64%  09-Aug-26
```












