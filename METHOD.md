# How this rebuild is done

The new payroll system is a new application. That is architectural work. The path below is the one we follow. A later step does not start until you accept the document from the step before it.

Three installed skills apply:

- **Superpowers** sets the order: understand, choose an approach, write the design, write the plan, then build with tests.
- **Ponytail** keeps the build small. Reuse what already exists. Do not add a feature the specification does not require.
- **Caveman** shortens conversation when you ask for it. It does not change the payroll rules or the code.

The old Visual Basic project and `HRAndPayroll_ReverseEngineering` stay read-only until a migration is its own accepted task.

## Order

| Step | What it produces | Status |
|---|---|---|
| 1. Read the old system | Notes from `TASAdmin.vbp` and the restored database | Done |
| 2. Lock the pay rules | `PAYROLL-SPEC.md` | Written. Waiting on your acceptance |
| 3. Choose the stack | Two or three options, one recommendation, your choice | Not started |
| 4. Design the new system | A design document: tables, screens, and how a cutoff runs. Reviewed in sections | Not started |
| 5. Plan the build | Small tasks. Each task names the files, the test, and the rule from `PAYROLL-SPEC.md` | Not started |
| 6. Build | One task at a time. A failing test first, then the smallest code that passes, then a review against the plan | Not started |

Step 4 is the database design. It comes after the stack, because keys and types follow the database you choose. Step 5 is the implementation plan. Code starts at step 6.

## Rules for every step

- `PAYROLL-SPEC.md` is the pay behavior. A design or a task that contradicts it stops until you change the specification.
- Staff time records and driver time records stay separate. Driver commission stays a trip import, not an hours calculation.
- Defects listed at the end of `PAYROLL-SPEC.md` are copied on purpose until you say to fix one.
- No Laravel, Vue, React, or other application code before step 6.
- No silent repair of swapped holiday names, forced 8-hour days, or the driver tax box.

## Where we are

Steps 1 and 2 are the reverse engineering. The next action is step 3: pick the stack. After that, the database design is step 4.
