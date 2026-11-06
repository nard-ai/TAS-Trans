# Server pay rules implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The locked pay rules run on the server, and tests fail if those rules drift.

**Architecture:** `src/server/engine.ts` stays the pay math. Screens keep calling `src/data/api.ts`, but that file only posts to `/api/call`. The Vite server loads the Neon database, runs the existing action, and saves the result. The browser does not import the engine.

**Tech Stack:** React, Vite, Neon Postgres, Node's built-in test runner.

**Spec:** `PAYROLL-SPEC.md`

## Global Constraints

- Staff required hours are forced to 8 after the punch math.
- Leave letter H is a special holiday and is paid with SpecialHolidayRate. Leave letter L is a legal holiday and is paid with LegalHolidayRate.
- If the cutoff end day is less than 17, staff payroll takes SSS and PhilHealth and skips Pag-IBIG. Otherwise it takes Pag-IBIG and skips SSS and PhilHealth. Withholding tax runs on both halves. If the end day is greater than 15, union dues are zero.
- When 13th-month pay is on, the cutoff amount is DailyRate / 12.
- Driver overtime in ProcessNewDTRCon stays 0. February stops at day 28. Driver tax stays the 200.00 box. Do not fix those until a later task says so.
- Do not modify the Visual Basic project or `HRAndPayroll_ReverseEngineering`.
- Do not import the old SQL Server rows.
- Do not commit unless asked.

## Review Focus

- A monthly employee's 13th-month line must stay DailyRate / 12, not monthly rate divided by something else. Task 1 pins the hourly DailyRate path. Task 4 pins the monthly employee.
- A cutoff ending on day 16 takes SSS and PhilHealth (day is still under 17) and already has union dues at zero (day is over 15). Task 1 tests day 15 and day 30. Task 4 adds day 16.
- Night differential is whole hours from 22:00 to 06:00 only when the shift night-diff flag is on. Task 3.
- Driver commission of 0 skips the contribution update, and a ceiling cuts the commission before the brackets. Task 3.
- Employees with payroll status T, W, X, Y, or Z stay off the payroll list. Task 4.

## File map

- `TAS Payroll Desk/src/server/engine.ts`: pay math. Already exists.
- `TAS Payroll Desk/src/server/actions.ts`: the current screen actions, run only on the server.
- `TAS Payroll Desk/src/server/desk.ts`: the database object those actions mutate.
- `TAS Payroll Desk/src/data/api.ts`: browser wrappers. No engine import.
- `TAS Payroll Desk/src/data/money.ts`: rounding, ids, and empty pay inputs. Safe for the browser.
- `TAS Payroll Desk/src/server/pay.test.mjs`: the rule tests and the "browser does not import the engine" test.

### Task 1: Four rules on the server

**Files:**
- Test: `TAS Payroll Desk/src/server/pay.test.mjs`
- Modify: `TAS Payroll Desk/src/server/engine.ts` (13th-month amount only, if the test fails)
- Create: `TAS Payroll Desk/src/server/actions.ts`, `TAS Payroll Desk/src/server/desk.ts`, `TAS Payroll Desk/src/data/money.ts`
- Modify: `TAS Payroll Desk/src/data/api.ts`, `TAS Payroll Desk/src/server/db-plugin.ts`, screen imports of `uid`, `r2`, and `emptyInputs`

**Interfaces:**
- Consumes: `staffBatch`, `attendanceTotals`, `computePayroll` from `engine.ts`
- Produces: `POST /api/call` with `{ name, args }` and a response `{ result, db }` or `{ error, field }`

- [x] **Step 1: Write the failing test**

Assert all of these:

- A staff day punched 08:00 to 12:00 has required hours 8 and work hours under 8.
- Attendance letter H sums into special hours and pays `hours * hourly * specialHoliday`. Letter L sums into legal hours and pays `hours * hourly * legalHoliday`.
- Cutoff end `2025-11-15` has deduction keys `sss`, `philhealth`, `tax`, and `union`, and no `pagibig`. End `2025-11-30` has `pagibig` and `tax`, and no `sss`, `philhealth`, or `union`.
- With 13th month on, daily rate 600, and 10 attendance days, the `thirteenth` earning amount is 50.
- No file under `src/screens`, `src/ui`, `src/app`, or `src/data/api.ts` contains `server/engine`.

- [x] **Step 2: Run the test and confirm the 13th-month assertion and the import assertion fail**

Run: `node --test src/server/pay.test.mjs` from `TAS Payroll Desk`.

- [x] **Step 3: Make the test pass**

Set the cutoff 13th-month earning to daily rate / 12. Move the action functions onto the server. Point `src/data/api.ts` at `POST /api/call`. Move `uid`, `r2`, and `emptyInputs` out of the engine file so screens can use them without importing the engine.

- [x] **Step 4: Run the test again**

Expected: pass. Then log in, open the work queue, and save a screen so a round trip still hits Neon.

### Task 2: Driver time defects stay

**Files:**
- Test: `TAS Payroll Desk/src/server/pay.test.mjs`
- Modify: none, unless a test shows the engine already repaired a defect

**Interfaces:**
- Consumes: `fmsBatch` from `engine.ts`

- [x] Assert driver overtime from the batch is 0 on a day whose pairs sum past 8.
- [x] Assert February does not create day 29.
- [x] Run the test. Do not change the engine if it already matches.

### Task 3: Driver commission defects stay

**Files:**
- Test: `TAS Payroll Desk/src/server/pay.test.mjs`
- Modify: `commissionContributions` only if it does not match the spec

**Interfaces:**
- Consumes: `commissionContributions` from `engine.ts`

- [x] Commission 0 does not produce a contribution update.
- [x] A ceiling of 20000 cuts a larger commission before the brackets.
- [x] The tax amount stored for the payment is 200, not a tax-table result.
- [x] SSS and PhilHealth use `minwage >= commission`. Pag-IBIG uses a normal between.

### Task 4: Payroll list and monthly 13th month

**Files:**
- Test: `TAS Payroll Desk/src/server/pay.test.mjs`
- Modify: the payroll list filter only if excluded statuses are included

**Interfaces:**
- Consumes: `isExcluded` from `src/data/rules.ts` and `computePayroll`

- [x] Statuses T, W, X, Y, and Z are excluded. A and P are not.
- [x] A monthly employee with daily rate 610 and 13th month on gets a thirteenth line of 610 / 12, rounded to cents.
- [x] Cutoff end day 16 includes SSS and PhilHealth, includes tax, and omits union dues and Pag-IBIG.
