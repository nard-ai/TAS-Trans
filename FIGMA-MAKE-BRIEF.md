# Figma Make brief: TAS payroll website

Paste this file into Figma Make. Design the interface. Do not add screens, buttons, or fields that are not listed here. Do not copy the sample layout that came with the color palette. That sample is only a color preview.

This is an internal HR and payroll office system for TAS Trans Group of Companies. It is not a store. There is no catalog, cart, checkout, or public signup.

The app is React. Data is stored in Neon (Postgres). A server computes pay. The screens collect input and show results. They do not calculate tax, SSS, PhilHealth, Pag-IBIG, night differential, or net pay in the browser.

## Visual limits

Use Inter or Poppins for all text. Use only the colors below.

Primary blue. Use for navigation, links, selected states, and emphasis.

| Hex | Role |
|---|---|
| #003E6B | Darkest blue |
| #0A558C | |
| #0F609B | |
| #186FAF | |
| #2680C2 | |
| #4098D7 | |
| #62B0E8 | |
| #84C5F4 | |
| #B6E0FE | |
| #DCEFFB | Lightest blue |

Yellow. Use for the main action on a screen (one primary button).

| Hex | Role |
|---|---|
| #8D2B0B | Darkest yellow-brown |
| #B44D12 | |
| #CB6E17 | |
| #DE911D | |
| #F0B429 | Main action |
| #F7C948 | |
| #FADB5F | |
| #FCE588 | |
| #FFF3C4 | |
| #FFFBEA | Lightest yellow |

Cyan. Use for success and positive status.

| Hex |
|---|
| #044E54 |
| #0A6C74 |
| #0E7C86 |
| #14919B |
| #2CB1BC |
| #38BEC9 |
| #54D1DB |
| #87EAF2 |
| #BEF8FD |
| #E0FCFF |

Red. Use for errors, destructive actions, and failed states.

| Hex |
|---|
| #610404 |
| #780A0A |
| #911111 |
| #A61B1B |
| #BA2525 |
| #D64545 |
| #E66A6A |
| #F29B9B |
| #FACDCD |
| #FFEEEE |

Blue grey. Use for text, backgrounds, borders, and secondary buttons.

| Hex |
|---|
| #102A43 |
| #243B53 |
| #334E68 |
| #486581 |
| #627D98 |
| #829AB1 |
| #9FB3C8 |
| #BCCCDC |
| #D9E2EC |
| #F0F4F8 |

## Product rules the screens must show

- A client is either staff or FMS. FMS means drivers and conductors. `ClientCode2 = FMS` is the FMS client. Everyone else is staff.
- Staff and FMS share payroll entry. They do not share time keeping or commissions.
- A cutoff has two halves. Half 1 key is `YYYYMM01`. Half 2 key is `YYYYMM02`.
- If the cutoff end day is before the 17th, payroll shows SSS and PhilHealth and hides Pag-IBIG. Otherwise it shows Pag-IBIG and hides SSS and PhilHealth. Tax shows on both halves. Union dues are zero when the end day is after the 15th.
- Staff hours worked on the payroll screen are the required hours from the time-record batch (forced to 8 per processed day). The manual time editor is not the pay rule.
- Leave letter H is special holiday. Leave letter L is legal holiday. Label them that way on every screen.
- A normal cutoff can add 13th month as daily rate divided by 12 when the client flag is on. The year-end 13th-month payslip prints the stored year amount. It does not divide the rate again.
- Driver money on the commission screen comes from trips in a separate FMS database, not from hours.
- Employees whose payroll status is T, W, X, Y, or Z do not appear in the payroll list.
- After payroll is moved to history, that cutoff is read-only. Working rows are gone. History remains.
- Do not design a billing screen.

## Who can see what

Account types that may run payroll: Accounting, Administrator, FMS Payroll, FMS Admin, Staff Admin, Staff Payroll.

An account whose type starts with FMS sees only FMS clients. An account whose type starts with STA sees only staff clients. Administrator and Accounting see the clients their rights allow.

A superuser sees every area. Everyone else sees an area only when that right is on:

| Right | Areas |
|---|---|
| Payroll | Payroll processing, payroll master, payslips |
| Employee information | Human resources |
| Time keeping | Import DTR, time keeping, holidays, leave, deductions |
| Client master | Client rates |
| Employee master | Employee payroll profile |
| User accounts | Users |
| Administrative | Company, shift, user history |
| Reports | HR reports, SSS, PhilHealth, HDMF |
| Maintenance | Banks, branches, departments, tax and contribution tables |

Add, Edit, and Delete are separate. Hide a button when that right is off.

Shell after login, on every screen except login:

- Company name: TAS Trans Group of Companies
- Product name: HR and Payroll
- Signed-in user name
- Log out
- Change password
- Only the menus this user is allowed to open

Home after login is a work queue, not a chart dashboard. Show the open cutoff (client, start, end, half, billed or open), how many employees are still unprocessed, and shortcuts into time keeping, payroll entry, and driver commissions when the user has those rights.

## 1. Login

Purpose: prove the user and open the shell.

Fields:

- Agency (company). Choosing it shows the agency name.
- Username
- Password (masked)
- Save password. Default off. Leave it off in the new design unless the user turns it on.

Buttons:

- Ok. Submits the form. Also submit when the user presses Enter in the password field.
- Cancel. Clears the form and stays on login.
- Deactivate user. Hidden until the account is already logged in somewhere else.

States:

- Wrong username or password: "Login failed." Stay on the form. Focus username. After more than 3 failures, show "Too many Invalid Logon Attempts" and end the session. There is no public sign-up and no forgot-password flow in this version.
- Correct password, but the account is already active: tell the user someone is still logged in, and reveal Deactivate user. Deactivate clears that lock and lets this login continue. Do not open the shell until deactivate succeeds or the other session is gone.
- Correct password and the account is free: open Home. Record the user on the shell. Record login time.
- Missing agency or database: block Ok and send the user to a database setup message. Do not invent a connection form with extra fields.

Do not show DTR, leave, or deduction shortcuts on the login screen. Those open only after a successful login.

## 2. Log out and change password

Log out: confirm, clear the active-login lock, record logout time, return to Login.

Change password: current password, new password, confirm new password. Save. Wrong current password stays on the form with an error. Success returns to the previous screen.

## 3. Company

Purpose: the agency printed on reports. This record can be empty today, so the screen must have an empty state that asks for the company name before reports can show a header.

Fields: company name, address, and the other identity fields already used on report headers. Save. Cancel.

Buttons: Save, Cancel.

Empty: explain that report headers stay blank until this is saved.

## 4. Clients and rate rules

Purpose: one row per client. This is where staff versus FMS, rates, and the 2018 tax switch live.

List: client code, name, status (active or not), staff or FMS. Search by name. Filter active only.

Open a client to edit:

- Name, code, active flag
- Staff or FMS (`FMS` or not)
- Working days per year (blank means 303 on the payroll screen)
- Daily hours basis is 8. Do not ask the user to invent another day length.
- Rates: overtime, night differential, rest day, special holiday, legal holiday, holiday overtime, excess overtime, ECOLA, hazard, agency rate, daily charge
- Flags: SSS, PhilHealth, Pag-IBIG, 13th month on the cutoff, display 13th month, ECOLA on rest day
- The ECOLA-on-rest-day checkbox is also the switch that uses the 2018 tax table. Label it "Use 2018 tax table" and keep the ECOLA behavior in the help text. Do not create two checkboxes.
- Tax percentage flag and the daily-rate ceiling under which tax is zero

Buttons: New client, Save, Cancel, Deactivate (sets inactive; does not delete history).

FMS accounts cannot open a staff client. Staff accounts cannot open an FMS client.

## 5. Employees

Purpose: the 201 file and the pay profile used by payroll.

List: employee number, name, client, position, payroll status, active. Search by name or number. Filter by client. People with payroll status T, W, X, Y, or Z can be stored here but are marked excluded from payroll.

Profile fields:

- Employee number (the key people see), internal id, DTR id (this is what time records join on)
- Name, department, position, tax code, civil status and the other HR facts already on the employee record
- Payment basis: Monthly or Hourly
- Monthly rate, daily rate, hourly rate
- SSS, PhilHealth, and Pag-IBIG on or off for this person
- PhilHealth amount per day when it replaces the table. Zero means use the table rule.
- Pag-IBIG amount per day when it replaces the table. Zero means use the table rule.
- Shift
- Client
- Hired, contract start, contract end

Buttons: New, Save, Cancel, Copy employee (copies the profile onto a new employee number the user types), Filter.

Human resources sections, each a list plus add, edit, and delete when the user has those rights:

- Benefits
- Assignments
- Trainings
- Offenses
- Leaves (balances and filed leave)
- Appraisals

Employee details is the read-only summary of the same person.

## 6. Shift, holiday, leave calendar

Shift schedule. One row per shift. Fields the time-record batch actually reads: code, start, end, midday, work hours, grace minutes, grace on or off, no late, no undertime, undertime minutes, no in/out, flexible time, no overtime, overtime start, night differential on or off. Buttons: New, Save, Delete.

Holiday table. Date, name, and letter. H is special holiday. L is legal holiday. Buttons: New, Save, Delete.

Leave and day off. Per employee, per date: letter and hours. Letters the payroll screen understands:

| Letter | Screen label | How payroll reads it |
|---|---|---|
| H | Special holiday | Hours, special-holiday rate |
| L | Legal holiday | Hours, legal-holiday rate |
| M, S | Sick leave | Hours divided by 8, as days |
| P, V, Y | Vacation leave | Hours divided by 8, as days |
| R | Rest day | Hours, and those hours are subtracted from hours worked |
| D | Day off worked | Rest-day overtime |
| B | Official business | Required hours become the leave hours |

Also file overtime for a date. That overtime replaces the computed overtime for that day.

Buttons: Save leave, Delete leave, Save overtime, Print leave balance.

## 7. Cutoff

Purpose: pick the period before time keeping or payroll.

Fields: client, start date, end date. The end day decides the half and which contributions appear.

Show: half (1 or 2), DTR key, whether this cutoff is already billed.

Buttons: Select cutoff (required before time keeping or payroll; both screens read this selection), Mark billed (only as the result of Move to master, not as a casual toggle).

If no cutoff is selected, time keeping and payroll entry show an error and do not open their grids.

## 8. Import time records

Purpose: bring punches in before the batch runs.

Fields: file or source of punches, DTR id, date, time, in or out. Show a preview list before saving.

Buttons: Preview, Import, Replace DTR id (old invalid id, new id, who changed it). Replace asks for confirmation. If the new id cannot be saved, the old rows stay.

Empty: no punches for this cutoff yet.

## 9. Staff time keeping

Shown only for a staff client.

List: employees for the client. Open one employee for the selected cutoff.

Day grid for each included day: date, weekday, time in, time out, required hours, work hours, late, undertime, overtime, night differential, leave letter, leave hours, remark (Present, Absent, No In/Out, Invalid Day).

Buttons:

- Compute. Runs the batch for this employee and half. Recompute asks for confirmation because it deletes the current final rows for that key first.
- Finalize for payroll. Copies the included final rows into the payroll attendance table. This is the set payroll will read.
- Save manual time. Edits one punch pair. Label it as a manual edit. Do not present it as the rule payroll uses after Compute.
- Refresh from leave and overtime. Updates leave without a full recompute when final rows already exist.

States:

- No punches: grid shows absent on included days that are not a day off.
- Equal time in and time out: remark No In/Out, hours 0.
- After finalize: payroll entry can load this employee. If finalize was skipped, payroll says no time records were processed.

Show a short note on the screen: required hours become 8 after compute, and that figure is what payroll pays.

## 10. Driver and conductor time keeping

Shown only for an FMS client. Same cutoff picker. Different grid.

Day grid: up to six in and out pairs, hours per pair, required hours as the sum of the pairs, remark. Do not show a forced 8. Do not show night differential or late as if the batch computed them. Those stay blank unless a later leave update sets them.

February days after the 28th and short-month days after the 30th show Invalid Day.

Buttons: Compute, Finalize for payroll, Refresh leave and overtime.

Note on the screen: overtime from this compute is 0. Rest-day overtime is applied on the last day only, and only for leave letter D.

## 11. Employee deductions

Purpose: loans and recurring deductions that payroll will take.

List by employee: code, name, balance, schedule, amount this cutoff, locked flag.

Codes and the labels to show:

| Code | Label |
|---|---|
| 1 | Uniform |
| 3 | Cash advance |
| 4 | Planters |
| 5 | SSS loan |
| 6 | Pag-IBIG loan |
| 7 | Vale |
| 8 | Other |
| 10, 11 | LMC |
| 12 | Surety or calamity |
| 16 | HDMF calamity |
| 17 | SSS calamity |

On an FMS payroll entry the labels change for the same slots: damage and accident, shortages, surety fund. Use those FMS labels only when the open client is FMS.

Buttons: New deduction, Save, Delete. If the row is locked, Save and Delete are disabled.

## 12. Payroll entry

Purpose: one employee, one cutoff, then save. This is the main work screen.

Header: client, cutoff start and end, half, DTR key, staff or FMS, user.

Employee list for the client: name, number, processed or not. Exclude payroll status T, W, X, Y, and Z. Count processed and remaining.

Select an employee. Show pay basis and rates (monthly, daily, or hourly, and tax code).

Load attendance button. If none was finalized, stop with "No DTR records processed" and the cutoff dates. Do not invent zeros and continue.

Hours and days the screen shows after load:

- Hours worked (staff: sum of required hours; monthly employees show 104)
- Days with required hours not zero
- Regular overtime hours
- Night differential hours
- Rest-day hours
- Special holiday hours (letter H)
- Legal holiday hours (letter L)
- Sick leave days
- Vacation leave days
- Absence and late amounts

Earnings the user can see, each with hours or quantity and the amount the server returns:

- Basic pay
- Overtime, night differential, rest day, special holiday, legal holiday
- ECOLA, hazard, allowance, other income, tax refund
- 13th month for this cutoff, only when the client flag is on. Label the amount as daily rate / 12. Do not also show the year-end figure here.

Deductions:

- First half: SSS employee, PhilHealth employee, tax. Pag-IBIG hidden.
- Second half: Pag-IBIG employee, tax. SSS and PhilHealth hidden.
- Union dues hidden when the end day is after the 15th
- Loans and other deductions from section 11
- Absence, late

Totals: gross, deductions, net pay. Amounts are formatted money. The server fills them. Fields that are results are read-only. Hours stay editable only before save, and saving sends them back to the server to recompute. Do not show a calculator keypad.

Buttons:

- Load attendance
- Recompute (asks the server to refill amounts)
- Save. Writes the open payroll for this employee and cutoff. Marks the employee processed.
- Move to master. Confirmation: this closes the cutoff for the whole client, copies open payroll into history, removes the open rows, and marks the cutoff billed. Disabled until the user confirms. Disabled if the cutoff is already billed.
- Print payslip. Disabled until this employee is saved.
- Back to employee list

If the cutoff is billed, every field is read-only and Save is hidden. Move to master is hidden. Print payslip remains.

FMS payroll entry uses the driver attendance totals and the FMS deduction labels. It does not import trips. Trip commissions stay on the next screen.

## 13. Payslip

Purpose: print what was saved. It does not recompute.

Show: employee, client, cutoff, each earning, each deduction, gross, net, days as days present divided by 8.

Staff labels include planters, other deductions, tax refund. FMS labels include damages, shortages, bonus, calamity loan.

Buttons: Print, Close.

Empty: no saved payroll for this employee and cutoff.

## 14. Payroll history

Purpose: closed payroll after Move to master.

List: client, cutoff, employee, gross, net, who processed it. Open is read-only. No Save. Print payslip works from history.

## 15. Driver commissions

Purpose: FMS only. Separate from payroll entry.

Fields: driver or conductor, employee, department, year, month, FMS database name, tax code.

Buttons:

- Load trips. Fills a list: date, bundle, commission, fund, bus, route. Empty list means that database has no trips for that person and month.
- Totals: commission, surety fund, ceiling, beginning balance. The tax box starts at 200.00 and the user can change it. Label it "Tax amount". Do not label it as a computed tax.
- Save payment. Replaces any payment already stored for that employee, month, and year. Then the server writes SSS, PhilHealth, Pag-IBIG, and tax onto that payment.
- If commission is 0, show that contributions were not updated.

Result row: employee share and employer share for SSS, PhilHealth, and Pag-IBIG, plus ECC and the tax amount from the box.

Note on the screen, in plain language: commission above the ceiling is cut to the ceiling before contributions. Tax is the amount in the tax box.

List of saved payments for the employee, filterable by quarter and year.

## 16. 13th-month payslip

Purpose: year-end print from the stored 13th-month amount.

Fields: client, year, employee. Show the stored amount as both gross and take-home. Show the pay period dates the user enters for the printout.

Buttons: Generate, Print.

Do not show daily rate / 12 on this screen. That figure belongs only on payroll entry.

## 17. Reports

Each report: filters (client, cutoff or year, employee when needed), a result table, Print, Export.

- Payroll summary for a cutoff
- SSS report
- PhilHealth report
- HDMF (Pag-IBIG) report
- HR reports: hired list, end of contract list
- User transaction history: who changed what, and when. Read-only.

Do not include billing reports.

## 18. Maintenance tables

Simple list plus new, save, and delete. Used by the server. The screens edit rows. They do not preview a whole payroll.

- Bank, branch, department
- SSS table: wage from, wage to, employee share, employer share, EC share
- PhilHealth table: wage from, wage to, employee share, employer share
- Pag-IBIG table: wage from, wage to, employee rate, employer rate
- Tax table: semi-monthly rows and the 2018 rows (bracket floor, percent, base tax, tax code)
- Tax exemption codes

## 19. Users

List: username, account type, active login, superuser.

Fields: username, password, account type, superuser, and one checkbox per right in the table in "Who can see what", plus Add, Edit, and Delete.

Buttons: New, Save, Deactivate login lock.

Password is entered, not shown back in the list.

## Buttons used across the product

Use these labels. Do not rename them per screen.

| Label | When it shows | Result |
|---|---|---|
| Ok | Login | Try to enter |
| Cancel | Forms | Discard unsaved edits and stay |
| Save | Any editor | Persist the open record |
| New | Lists | Empty form |
| Delete | A saved row, and only with delete right | Confirm, then remove |
| Deactivate user | Login, only if already active | Clear the lock |
| Log out | Shell | End the session |
| Select cutoff | Cutoff | Store the period for the next screens |
| Preview | Import | Show punches without saving |
| Import | Import | Save punches |
| Compute | Time keeping | Run the batch |
| Finalize for payroll | Time keeping | Copy attendance into the payroll set |
| Load attendance | Payroll entry | Pull finalized time records |
| Recompute | Payroll entry | Ask the server for amounts again |
| Move to master | Payroll entry | Close the client cutoff into history |
| Print | Payslip, reports | Print the saved result |
| Load trips | Driver commissions | Read the FMS trip database |
| Generate | 13th month | Fill the printout from the stored amount |

Destructive buttons (Delete, Move to master, Replace DTR id) ask for confirmation and use the red family. The single main action on a screen uses the yellow family. Every other action is secondary, in blue grey or blue.

## Empty, error, and saved

Every list has an empty state with one sentence and the button that creates the first row.

Every failed save names the field or the rule (no cutoff, no finalized time records, already billed, no tax code, tax code missing from the table, commission is zero).

Every successful save names what was saved (employee, cutoff, or payment) and stays on the record.

## Out of the design

- Ecommerce, marketing pages, and a chart-only dashboard
- Billing
- A second set of holiday names that calls H legal or L special
- Pay math drawn as a spreadsheet the user fills by hand
- Colors outside the palette, and fonts other than Inter or Poppins
