# Payroll specification

Locked from the Visual Basic 6 program `TASAdmin.vbp` and the restored database `HRAndPayroll_ReverseEngineering`. This document records the rules the new system must follow. It does not change the old program or the database.

A client is an FMS (driver and conductor) client when `ClientTable.ClientCode2` is `FMS`. Every other client is staff. Staff and FMS share the payroll screen. They do not share the time-record procedure or the commission screen.

## Decisions

1. Staff daily time records follow `ProcessNewDTRCutoff`. Grace is the shift `GracePeriod`. Night differential is whole hours from 22:00 to 06:00, and only when `HRShiftSchedule.NoNightDiff` is on. Required hours are forced to 8. The time-record screen's fixed 10-minute grace is a manual editor, not the pay rule.
2. On a normal cutoff, when 13th-month pay is on, the amount is `DailyRate / 12`. The year-end payslip prints `HR13thMonthPay.month13payment` and does not divide the rate again.
3. Leave letter `H` is a special holiday and is paid with `SpecialHolidayRate`. Leave letter `L` is a legal holiday and is paid with `LegalHolidayRate`. The old textbox names are reversed and are not the names to keep.
4. Driver and conductor payroll is in scope. Time records follow `ComputeDTRCon`. Commissions follow `frmFMSDeduction`, `GetFMSCommissions`, and `DriverConductorSSSComputation`.

## Cutoff

`DTRCutoffTable` holds one row of first-half and second-half dates. Half 1 builds `DTRKey` as `YYYYMM01`. Half 2 builds `YYYYMM02`.

If the cutoff end day is less than 17, the staff payroll screen takes SSS and PhilHealth and skips Pag-IBIG. Otherwise it takes Pag-IBIG and skips SSS and PhilHealth. Withholding tax runs on both halves. If the end day is greater than 15, union dues are zero.

`frmPayrollEntry` loads attendance, computes pay, and writes `PayrollTrans` and `PayrollAdjust`. Move to master copies those rows into `PayrollMaster` and `MasterAdjust` for that client, deletes the working rows, and marks `CutoffTable` billed. There is no stored procedure for that copy.

Employees with `payrollstatus` in `T`, `W`, `X`, `Y`, or `Z` are left out of the payroll list. The business meaning of those letters is not in the code.

## Staff time records

`frmTimeKeep` calls `ComputeDTRCutOff`. Mode 1 computes. Mode 2 marks `DTRFinalRecords` final and copies included rows into `DTRPayrollRecords`.

`ProcessNewDTRCutoff` writes scratch table `DTRNoOfDays`, then inserts `DTRFinalRecords`.

- Included days that are not a day off start as absent.
- `GetFiledLeaveAndOTCutOff` copies `HRDayoffCalendar` and `HRFiledOvertime` onto those days.
- Time in is the earliest punch. Time out is the latest punch. A shift that crosses midnight uses `_MidnightTime`.
- Shift flags are read once, from the first punch's shift, and reused for every day in the run.
- Blank midday becomes 12:00. Zero shift hours become 8. Break is 1 hour.
- `NoInOut` replaces both punches with the shift. `NoLate` on a fixed shift replaces time in with the shift start.
- A punch pair that ends before midday is a half day. Flexible time uses `_GetFlexiHours` and clears late, undertime, and overtime. A fixed shift sets work time to `8` minus late minutes from `_GetWorkHours`.
- Late and undertime use `_GetWorkHours` in minutes. Grace minutes are subtracted only when the grace flag is on and the grace amount is not zero.
- Equal time in and time out is `No In/Out`: work time 0, late 0, undertime 0.
- Night differential uses `_GetNightDiffHours`, then drops the fraction, then caps at 8.
- After that math, required hours are set to 8 even when the punch span is shorter. Payroll reads the sum of required hours as hours worked (`vwDTRFinalPayroll.WorkHours`).
- Leave types `D`, `V`, and `S` clear times, required hours, work hours, and overtime. Types `B` and `R` set required hours to the leave hours. Type `D` also sets rest-day overtime to elapsed hours minus 1. Filed overtime overwrites `DTROT`.
- `vwFiledLeave` divides hours by 8 for letters `P`, `V`, `Y`, `M`, and `S`. Letters `H`, `L`, and `R` stay in hours.

`GetDTRRecords` and `ReplaceDTRID` are hard-coded to database name `HRAndPayroll`. On this restored copy that name does not exist.

## Staff pay

Rates come from the employee pay profile and `ClientTable`. If `WorkingDaysperYear` is null, use 303. A monthly employee is forced to 104 hours on the screen, and basic pay for the cutoff is `MonthlyRate / 2`. Hourly amounts use `((MonthlyRate * 12) / WorkingDaysperYear) / 8`. ECOLA uses the same divisor.

Holiday hours:

| Letter | Meaning | Rate | Stored hours | Stored pay |
|---|---|---|---|---|
| H | Special holiday | SpecialHolidayRate | NoSplHoliday | SplHolidayPay |
| L | Legal holiday | LegalHolidayRate | NoRegHoliday | RegHolidayPay |

Overtime, night differential, and rest day multiply the hour amount by the client rate. A zero rest-day rate is treated as 1 in `ComputeAmount`, which means the premium is not added on top of a zero. The screen still has default premiums of 0.3, 1.5, and 2.6 in other branches when the client rate is 0. Those defaults stay where that code applies. Do not invent a single default for every premium.

Gross pay subtracts absence and late amounts from the sum of earning boxes.

SSS uses `SSSTable` on current gross plus the previous `PayrollMaster.GrossPay` for the other half. Employee share is `ee_share`. Employer share is `er_share` plus `ec_share`.

PhilHealth uses `NewPHealthRates` on basic pay. If the client PhilHealth rate is not 0, use that rate for both shares. If basic pay is 10,000 or less, each share is 150. Otherwise the premium is `basic * (NewPHRate / 100)`. If the last cent of that premium is odd, the employee share is one cent below half and the employer share is half. Otherwise each share is half.

Pag-IBIG uses `PagibigTable`. The employee share is `basic * pagibigee`. The employer share uses `pagibiger`. A per-day Pag-IBIG rate on the employee, when not zero, replaces the table.

Tax uses `SemiMonthlyTax2018` when `ClientTable.EcolaPayonRestDay` is true. That column is the 2018-tax switch, not only an ECOLA flag. Otherwise tax uses `SemiMonthlyTaxTable` filtered by the employee tax code. The 2018 path does not filter by tax code. The amount is `((taxable - SalRate2) * PerOver) + BaseTax` from the bracket the code walks to. If the client tax-percentage flag is on and the daily rate is at or below `TaxPercentage`, tax is zero.

When 13th-month pay is on, the cutoff amount is `DailyRate / 12`, including monthly employees. The year-end payslip is a separate print of `HR13thMonthPay.month13payment` through `Generate13thMonthPayslip3`.

Loans and other deductions come from `EmployeeDeductions` and `EmployeeDedSchedules`. The payroll screen maps deduction codes onto reused columns. The names to keep are the screen labels, not the old column names: uniform, cash advance, planters, SSS loan, Pag-IBIG loan, vale, other, LMC, surety or calamity, HDMF calamity, SSS calamity. `GetPayrollEntries` only copies a saved payroll into `tmpPayrollTemplate` for the payslip. It does not compute pay. For FMS it relabels damages, shortages, bonus, and calamity loan on that template.

## Driver and conductor time records

`frmTimeKeep` calls `ComputeDTRCon` with the same year, month, half, DTR id, mode, and reprocess flag as the staff call. Mode 2 copies `DTRFinalConductors` into `DTRPayrollConductors`. Mode 1 calls `ProcessNewDTRCon`, or `UpdateFiledLeaveAndOTCon` when rows already exist and reprocess is off.

`vwDTRConductors` reads `DTRRecords`. Each punch is one row. Time in and time out on that row are the same punch, adjusted by `_MidnightTime`. `DTR_IOMode` says whether the punch is an out. The employee's own shift supplies the schedule. A missing shift becomes shift 10, 06:00 to 14:00.

`ProcessNewDTRCon` writes scratch table `DTRNoOFDaysCon`, then inserts `DTRFinalConductors`.

- February days after 28 are `Invalid Day`. April, June, September, and November days after 30 are `Invalid Day`. Day 29 of a leap year is not created.
- Up to six in/out pairs are stored (`TimeIN1` through `TimeIN6`).
- Each pair's hours are `_WorkHours` of that pair. Required hours are the running sum of those pairs. They are not forced to 8.
- The day-close overtime block sets work time to 8 when the sum is above 8, then sets overtime to `CONVERT(int, 8 - 8)`, which is 0. Overtime from this block is always 0.
- Rest-day overtime inside the day loop is commented out. It runs only for the last day of the cursor, and only for leave type `D`. Official business (type `B`) on that last day clears rest-day overtime, overtime, and late.
- Late, undertime, grace, flexible time, and night differential are selected from the shift and then not written. The payroll view still sums `DTRLates`, `DTRUnderTime`, `NDiff`, and `DTROT`, which stay 0 unless leave updates them.
- Absent is not set. The absent update is commented out.
- `vwDTRFinalConductor` sums required hours as `WorkHours` for `frmPayrollEntry`. The same leave-letter mapping as staff applies, from `vwFiledLeave` when that view has rows for the DTR id. Days present are the count of conductor rows whose required hours are not 0.

`GetDTRRecordsCon` is hard-coded to `HRAndPayroll..vwDTRReportCon`.

## Driver and conductor commissions

Commissions are not computed from hours. `frmFMSDeduction` calls `GetFMSCommissions`.

The procedure truncates `tmpFMSCommissions` and inserts from `@DBName..Trips` for the year and the month name. The month argument is the combo text with the first three characters removed (`01 January` becomes `January`). A driver uses `ID_DRV`, `D_Comission`, and `D_Fund`. A conductor uses `ID_CON`, `C_Comission`, and `C_Fund`. The department filter is commented out. Trip rows are not in `HRAndPayroll`. They live in the FMS database named on the screen.

Saving a payment deletes any `DriverConductorPayments` row for that employee, applied month, and applied year, then inserts client, employee, position (`DRIVER` or conductor), department, tax code, beginning balance, commission, and surety fund.

`DriverConductorSSSComputation` then updates that row. Parameters from the screen, in order: ceiling (`txtFund(3)`), commission (`txtFund(1)`), surety fund (`txtFund(2)`), month combo, year, name, employee id, tax code, beginning balance, FMS database name, department, `txtFund(0)`, user id, and `txtFund(0)` again. `txtFund(0)` is initialized to `200.00`. The bundle count is calculated and not written into that box.

Rules inside the procedure:

- Commission of 0 skips the update.
- If the ceiling is not 0 and commission is above the ceiling, commission is cut to the ceiling before the brackets.
- SSS and PhilHealth walk `HRANDPayroll..SSSTable` and `HRANDPayroll..PhilHealthTable` with `minwage >= commission AND commission <= maxwage`. That is the test in the procedure. Pag-IBIG walks `HRANDPayroll..PagibigTable` with `commission BETWEEN minwage AND maxwage`, then sets the employee share to `commission * PagibigEE` and the employer share to `commission * PagibigER`.
- The monthly tax cursor runs and is discarded. `EETax` is the `txtFund(0)` value, not `((commission - SalRate2) * PerOver) + BaseTax`.
- The procedure is hard-coded to database name `HRANDPayroll`.

Those statutory amounts are driver commission deductions. They are not the staff cutoff split in `frmPayrollEntry`. An FMS employee who is also run through the payroll screen still uses the staff pay formulas, with attendance from `vwDTRFinalConductor` and the FMS labels for damage, shortages, and surety fund.

## Tables the new system must represent

People and rules: `CompanyTable`, `ClientTable`, `HREmployees`, `EmployeeBasicSalary`, `HRShiftSchedule`, `HRLeaveType`, `HRDayoffCalendar`, `HRFiledOvertime`, `UsersList`.

Staff attendance: `DTRRecords`, `DTRNoOfDays`, `DTRFinalRecords`, `DTRPayrollRecords`, `DTRCutoffTable`.

Driver attendance: `DTRNoOFDaysCon`, `DTRFinalConductors`, `DTRPayrollConductors`.

Pay: `CutoffTable`, `PayrollTrans`, `PayrollAdjust`, `PayrollMaster`, `MasterAdjust`, `EmployeeDeductions`, `EmployeeDedSchedules`, `HR13thMonthPay`, `tmpPayrollTemplate`.

Statutory tables: `SSSTable`, `PhilhealthTable`, `PagibigTable`, `SemiMonthlyTaxTable`, `semimonthlytax2018`, `MonthlyTaxTable`, `TaxExemptionCode`.

Driver pay: `DriverConductorPayments`, `tmpFMSCommissions`, plus the external `Trips` table.

`HREmployees` is keyed on `EmpNo`. Payroll joins attendance on `DTRID`. There are no foreign keys in the restored database. Dated copies such as `DTRRecordsFeb2016` are snapshots, not the live path.

## Left out

- The billing forms. They are not loaded by `TASAdmin.vbp`.
- SQL source-control objects (`dt_*`, `dtproperties`).
- Rebuilding the manual time editor's 10-minute grace as the pay rule.
- A new web stack, a new database, and a data migration. Those start only after this specification is accepted.
- Silent repairs of the defects below.

## Defects to keep until you say otherwise

These are what the current program does. The first rebuild matches them so pay stays the same.

- Staff required hours are forced to 8 after the punch math.
- Driver overtime in `ProcessNewDTRCon` becomes 0 because work time is capped at 8 before the subtraction.
- Driver February stops at day 28.
- Driver SSS and PhilHealth use `minwage >= commission`, not a normal bracket. Pag-IBIG uses `BETWEEN`.
- Driver tax ignores `MonthlyTaxTable` and stores `txtFund(0)` (default 200.00).
- `GetDTRRecords`, `GetDTRRecordsCon`, `ReplaceDTRID`, and `DriverConductorSSSComputation` name the database `HRAndPayroll` or `HRANDPayroll`.
- `CompanyTable` has no rows in this restored copy.
