# 📊 Employee Attendance Tracker with Automated Alerts (Excel)

An Excel-based MIS tool that automatically tracks employee attendance, calculates late arrivals, and flags employees with repeated lateness — built to demonstrate how routine HR/MIS reporting can be automated using formulas instead of manual review.

## 🔍 Overview

This tracker processes **440+ daily attendance records** for **20 employees over 22 working days** (June 2026). It automatically:

- Calculates how many minutes late each employee checked in (with a grace period)
- Flags each record as **"On Time"** or **"LATE ARRIVAL - X min"**
- Counts late days per employee and flags **"FREQUENT LATECOMER - ACTION REQUIRED"** when an employee is late 3 or more times
- Generates an overall summary with punctuality rate, total late arrivals, and average delay

No manual checking needed — everything updates automatically when new attendance data is entered.

## 📁 File

- `Attendance_Tracker_Large.xlsx` — the complete workbook with daily log, employee-wise summary, and overall summary

## 🖼️ Screenshots
1. Daily attendance log with conditional formatting (red = late, green = on time)
 <img width="1191" height="886" alt="image" src="https://github.com/user-attachments/assets/0f0e19a9-71a7-430b-bd45-a82c19197160" />

3. Employee-wise late summary with HR alerts
 <img width="992" height="778" alt="image" src="https://github.com/user-attachments/assets/242dbeef-ca9b-4661-878b-a41e63d31526" />

5. Overall summary KPIs
   <img width="700" height="383" alt="image" src="https://github.com/user-attachments/assets/b59e1db5-586b-43ed-b195-868013d193ed" />


## ⚙️ How It Works

### 1. Calculating Late Minutes
```excel
=MAX(0, ROUND((TIMEVALUE(CheckInTime) - TIMEVALUE(ShiftStart)) * 1440, 0) - 10)
```
- `TIMEVALUE()` converts time text (e.g., "09:48 AM") into Excel's internal time fraction
- Subtracting shift start from check-in time gives the time difference as a fraction of a day
- Multiplying by `1440` (minutes in a day) converts this into minutes
- Subtracting `10` applies a 10-minute grace period
- `MAX(0, ...)` ensures early arrivals never show as negative late minutes

### 2. Attendance Status Flag
```excel
=IF(LateMinutes = 0, "On Time", "LATE ARRIVAL - " & LateMinutes & " min")
```
Converts the numeric result into a readable status, which conditional formatting then highlights in red or green.

### 3. Counting Late Days per Employee
```excel
=SUMPRODUCT((EmpID_Range = CurrentEmpID) * (LEFT(Status_Range, 4) = "LATE"))
```
- Compares every row's employee ID to the current employee
- Checks if that row's status starts with "LATE"
- Multiplies the two TRUE/FALSE arrays (TRUE = 1, FALSE = 0) so only matching rows count
- `SUMPRODUCT` adds these up to get the total late days for that employee

### 4. Frequent Latecomer Alert
```excel
=IF(DaysLate >= 3, "FREQUENT LATECOMER - ACTION REQUIRED", "OK")
```
Flags employees crossing the threshold so HR can take action without scanning the entire log.

### 5. Overall Summary KPIs
Built using `COUNTIF`, `COUNTA`, and `AVERAGEIF` to calculate:
- Total attendance records
- Total late arrivals vs. on-time records
- Overall punctuality percentage
- Number of employees flagged
- Average late minutes (for late records only)

## 📈 Sample Results

| Metric                          | Value  |
|----------------------------------|--------|
| Total Attendance Records         | 440    |
| Total Late Arrivals              | 58     |
| Overall Punctuality Rate         | 86.8%  |
| Employees Flagged as Frequent Latecomers | 7      |
| Average Late Minutes (when late) | 19.4   |

## 🛠️ Skills Demonstrated

- Excel formulas: `IF`, `TIMEVALUE`, `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGEIF`
- Conditional formatting for automated visual alerts
- MIS-style report design and KPI summarization
- Data-driven decision support for HR/operations teams
---

*Built as part of an MIS Executive portfolio project to demonstrate automated operational reporting.*
