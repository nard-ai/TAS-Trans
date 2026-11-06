import fs from "node:fs";
import path from "node:path";
import { neon, type NeonQueryFunction } from "@neondatabase/serverless";
import type { Db, FmsDay, StaffDay } from "../data/types";

type Sql = NeonQueryFunction<false, false>;

const TABLES = [
  "dataset",
  "agencies",
  "users",
  "company",
  "clients",
  "shifts",
  "holidays",
  "employees",
  "leaves",
  "overtime_filings",
  "punches",
  "staff_days",
  "fms_days",
  "payroll_attendance",
  "cutoffs",
  "deductions",
  "payroll_records",
  "trips",
  "commission_payments",
  "thirteenth_month",
  "banks",
  "branches",
  "departments",
  "sss_table",
  "philhealth_table",
  "hdmf_table",
  "tax_table",
  "tax_codes",
  "user_logs",
  "benefits",
  "assignments",
  "trainings",
  "offenses",
  "leave_balances",
  "appraisals",
  "desk_session",
];

export function connect(url: string): Sql {
  const cleaned = new URL(url);
  cleaned.searchParams.delete("channel_binding");
  return neon(cleaned.toString());
}

export function splitSql(text: string): string[] {
  return text
    .split(/;\s*(?:\r?\n|$)/)
    .map((part) => part.replace(/--[^\n]*/g, "").trim())
    .filter((part) => part.length > 0);
}

export async function ensureSchema(sql: Sql, schemaPath: string) {
  const statements = splitSql(fs.readFileSync(schemaPath, "utf8"));
  for (const statement of statements) {
    await sql.query(statement);
  }
}

export function schemaPath(root: string) {
  return path.join(root, "src", "server", "schema.sql");
}

function parse<T>(value: unknown): T {
  if (typeof value === "string") return JSON.parse(value) as T;
  return value as T;
}

function docs<T>(rows: Record<string, unknown>[]): T[] {
  return rows.map((row) => parse<T>(row.doc));
}

function dayMap<T>(rows: Record<string, unknown>[], column: string): Record<string, T> {
  const out: Record<string, T> = {};
  for (const row of rows) out[`${row.emp_id}|${row.dtr_key}`] = parse<T>(row[column]);
  return out;
}

export async function loadDb(sql: Sql): Promise<Db | null> {
  const marker = await sql`SELECT 1 FROM dataset WHERE id = 1`;
  if (marker.length === 0) return null;

  const [
    agencies,
    users,
    company,
    clients,
    employees,
    shifts,
    holidays,
    leaves,
    overtimes,
    punches,
    staffFinal,
    fmsFinal,
    payrollAttendance,
    cutoffs,
    deductions,
    openPayroll,
    history,
    trips,
    payments,
    thirteenth,
    banks,
    branches,
    departments,
    sss,
    philhealth,
    hdmf,
    tax,
    taxCodes,
    logs,
    benefits,
    assignments,
    trainings,
    offenses,
    leaveBalances,
    appraisals,
    desk,
  ] = await Promise.all([
    sql`SELECT doc FROM agencies ORDER BY ord`,
    sql`SELECT doc FROM users ORDER BY ord`,
    sql`SELECT doc FROM company WHERE id = 1`,
    sql`SELECT doc FROM clients ORDER BY ord`,
    sql`SELECT doc FROM employees ORDER BY ord`,
    sql`SELECT doc FROM shifts ORDER BY ord`,
    sql`SELECT doc FROM holidays ORDER BY ord`,
    sql`SELECT doc FROM leaves ORDER BY ord`,
    sql`SELECT doc FROM overtime_filings ORDER BY ord`,
    sql`SELECT doc FROM punches ORDER BY ord`,
    sql`SELECT emp_id, dtr_key, days FROM staff_days`,
    sql`SELECT emp_id, dtr_key, days FROM fms_days`,
    sql`SELECT emp_id, dtr_key, doc FROM payroll_attendance`,
    sql`SELECT doc FROM cutoffs ORDER BY ord`,
    sql`SELECT doc FROM deductions ORDER BY ord`,
    sql`SELECT doc FROM payroll_records WHERE open = true ORDER BY ord`,
    sql`SELECT doc FROM payroll_records WHERE open = false ORDER BY ord`,
    sql`SELECT doc FROM trips ORDER BY ord`,
    sql`SELECT doc FROM commission_payments ORDER BY ord`,
    sql`SELECT doc FROM thirteenth_month ORDER BY ord`,
    sql`SELECT doc FROM banks ORDER BY ord`,
    sql`SELECT doc FROM branches ORDER BY ord`,
    sql`SELECT doc FROM departments ORDER BY ord`,
    sql`SELECT doc FROM sss_table ORDER BY ord`,
    sql`SELECT doc FROM philhealth_table ORDER BY ord`,
    sql`SELECT doc FROM hdmf_table ORDER BY ord`,
    sql`SELECT doc FROM tax_table ORDER BY ord`,
    sql`SELECT doc FROM tax_codes ORDER BY ord`,
    sql`SELECT doc FROM user_logs ORDER BY ord`,
    sql`SELECT doc FROM benefits ORDER BY ord`,
    sql`SELECT doc FROM assignments ORDER BY ord`,
    sql`SELECT doc FROM trainings ORDER BY ord`,
    sql`SELECT doc FROM offenses ORDER BY ord`,
    sql`SELECT doc FROM leave_balances ORDER BY ord`,
    sql`SELECT doc FROM appraisals ORDER BY ord`,
    sql`SELECT session, flash FROM desk_session WHERE id = 1`,
  ]);

  const sessionRow = desk[0];
  return {
    agencies: docs(agencies),
    users: docs(users),
    company: company[0] ? parse(company[0].doc) : null,
    clients: docs(clients),
    employees: docs(employees),
    shifts: docs(shifts),
    holidays: docs(holidays),
    leaves: docs(leaves),
    overtimes: docs(overtimes),
    punches: docs(punches),
    staffFinal: dayMap<StaffDay[]>(staffFinal, "days"),
    fmsFinal: dayMap<FmsDay[]>(fmsFinal, "days"),
    payrollAttendance: dayMap(payrollAttendance, "doc"),
    cutoffs: docs(cutoffs),
    deductions: docs(deductions),
    openPayroll: docs(openPayroll),
    history: docs(history),
    trips: docs(trips),
    payments: docs(payments),
    thirteenth: docs(thirteenth),
    banks: docs(banks),
    branches: docs(branches),
    departments: docs(departments),
    sss: docs(sss),
    philhealth: docs(philhealth),
    hdmf: docs(hdmf),
    tax: docs(tax),
    taxCodes: docs(taxCodes),
    logs: docs(logs),
    benefits: docs(benefits),
    assignments: docs(assignments),
    trainings: docs(trainings),
    offenses: docs(offenses),
    leaveBalances: docs(leaveBalances),
    appraisals: docs(appraisals),
    session: sessionRow ? parse(sessionRow.session) : null,
    flash: sessionRow ? parse(sessionRow.flash) : null,
  };
}

function rows(
  sql: Sql,
  table: string,
  columns: string,
  exprs: string,
  payload: unknown,
) {
  return sql.query(
    `INSERT INTO ${table} (${columns}ord, doc)
     SELECT ${exprs}x.ord::int, x.doc
     FROM jsonb_array_elements($1::jsonb) WITH ORDINALITY AS x(doc, ord)`,
    [JSON.stringify(payload)],
  );
}

function chunkRows(
  sql: Sql,
  table: string,
  columns: string,
  exprs: string,
  payload: unknown[],
  size = 400,
) {
  if (payload.length === 0) return [rows(sql, table, columns, exprs, [])];
  const queries = [];
  for (let i = 0; i < payload.length; i += size)
    queries.push(rows(sql, table, columns, exprs, payload.slice(i, i + size)));
  return queries;
}

export async function saveDb(sql: Sql, db: Db) {
  const company = db.company ? JSON.stringify(db.company) : "null";
  await sql.transaction([
    sql.query(`TRUNCATE TABLE ${TABLES.join(", ")} RESTART IDENTITY CASCADE`),
    sql.query(`INSERT INTO dataset (id, initialized) VALUES (1, true)`),
    rows(
      sql,
      "agencies",
      "code, name, database_name, ",
      "x.doc->>'code', x.doc->>'name', x.doc->>'database', ",
      db.agencies,
    ),
    rows(
      sql,
      "users",
      "id, username, name, account_type, active_login, ",
      "x.doc->>'id', x.doc->>'username', x.doc->>'name', x.doc->>'accountType', (x.doc->>'activeLogin')::boolean, ",
      db.users,
    ),
    sql.query(
      `INSERT INTO company (id, name, doc)
       SELECT 1, $1::jsonb->>'name', $1::jsonb
       WHERE $1::jsonb <> 'null'::jsonb`,
      [company],
    ),
    rows(
      sql,
      "clients",
      "id, code, code2, name, active, ",
      "x.doc->>'id', x.doc->>'code', x.doc->>'code2', x.doc->>'name', (x.doc->>'active')::boolean, ",
      db.clients,
    ),
    rows(sql, "shifts", "id, code, ", "x.doc->>'id', x.doc->>'code', ", db.shifts),
    rows(
      sql,
      "holidays",
      "id, holiday_date, name, letter, ",
      "x.doc->>'id', x.doc->>'date', x.doc->>'name', x.doc->>'letter', ",
      db.holidays,
    ),
    rows(
      sql,
      "employees",
      "id, emp_no, dtr_id, client_id, payroll_status, role, active, ",
      "x.doc->>'id', x.doc->>'empNo', x.doc->>'dtrId', x.doc->>'clientId', x.doc->>'payrollStatus', x.doc->>'role', (x.doc->>'active')::boolean, ",
      db.employees,
    ),
    rows(sql, "leaves", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.leaves),
    rows(sql, "overtime_filings", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.overtimes),
    rows(
      sql,
      "punches",
      "id, dtr_id, punch_date, ",
      "x.doc->>'id', x.doc->>'dtrId', x.doc->>'date', ",
      db.punches,
    ),
    sql.query(
      `INSERT INTO staff_days (emp_id, dtr_key, days)
       SELECT split_part(kv.key, '|', 1),
              substr(kv.key, strpos(kv.key, '|') + 1),
              kv.value
       FROM jsonb_each($1::jsonb) AS kv`,
      [JSON.stringify(db.staffFinal)],
    ),
    sql.query(
      `INSERT INTO fms_days (emp_id, dtr_key, days)
       SELECT split_part(kv.key, '|', 1),
              substr(kv.key, strpos(kv.key, '|') + 1),
              kv.value
       FROM jsonb_each($1::jsonb) AS kv`,
      [JSON.stringify(db.fmsFinal)],
    ),
    sql.query(
      `INSERT INTO payroll_attendance (emp_id, dtr_key, kind, finalized_at, doc)
       SELECT split_part(kv.key, '|', 1),
              substr(kv.key, strpos(kv.key, '|') + 1),
              kv.value->>'kind',
              kv.value->>'at',
              kv.value
       FROM jsonb_each($1::jsonb) AS kv`,
      [JSON.stringify(db.payrollAttendance)],
    ),
    rows(
      sql,
      "cutoffs",
      "client_id, dtr_key, start_date, end_date, billed, ",
      "x.doc->>'clientId', x.doc->>'key', x.doc->>'start', x.doc->>'end', (x.doc->>'billed')::boolean, ",
      db.cutoffs,
    ),
    rows(
      sql,
      "deductions",
      "id, emp_id, code, ",
      "x.doc->>'id', x.doc->>'empId', (x.doc->>'code')::int, ",
      db.deductions,
    ),
    rows(
      sql,
      "payroll_records",
      "id, emp_id, client_id, dtr_key, half, open, gross, net, ",
      "x.doc->>'id', x.doc->>'empId', x.doc->>'clientId', x.doc->>'key', (x.doc->>'half')::int, true, (x.doc->>'gross')::numeric, (x.doc->>'net')::numeric, ",
      db.openPayroll,
    ),
    ...chunkRows(
      sql,
      "payroll_records",
      "id, emp_id, client_id, dtr_key, half, open, gross, net, ",
      "x.doc->>'id', x.doc->>'empId', x.doc->>'clientId', x.doc->>'key', (x.doc->>'half')::int, false, (x.doc->>'gross')::numeric, (x.doc->>'net')::numeric, ",
      db.history,
    ),
    rows(
      sql,
      "trips",
      "id, emp_id, trip_date, ",
      "x.doc->>'id', x.doc->>'empId', x.doc->>'date', ",
      db.trips,
    ),
    rows(
      sql,
      "commission_payments",
      "id, emp_id, pay_year, pay_month, ",
      "x.doc->>'id', x.doc->>'empId', (x.doc->>'year')::int, (x.doc->>'month')::int, ",
      db.payments,
    ),
    rows(
      sql,
      "thirteenth_month",
      "id, emp_id, client_id, pay_year, ",
      "x.doc->>'id', x.doc->>'empId', x.doc->>'clientId', (x.doc->>'year')::int, ",
      db.thirteenth,
    ),
    rows(sql, "banks", "id, code, ", "x.doc->>'id', x.doc->>'code', ", db.banks),
    rows(sql, "branches", "id, code, ", "x.doc->>'id', x.doc->>'code', ", db.branches),
    rows(sql, "departments", "id, code, ", "x.doc->>'id', x.doc->>'code', ", db.departments),
    rows(
      sql,
      "sss_table",
      "id, wage_from, wage_to, ",
      "x.doc->>'id', (x.doc->>'from')::numeric, (x.doc->>'to')::numeric, ",
      db.sss,
    ),
    rows(
      sql,
      "philhealth_table",
      "id, wage_from, wage_to, ",
      "x.doc->>'id', (x.doc->>'from')::numeric, (x.doc->>'to')::numeric, ",
      db.philhealth,
    ),
    rows(
      sql,
      "hdmf_table",
      "id, wage_from, wage_to, ",
      "x.doc->>'id', (x.doc->>'from')::numeric, (x.doc->>'to')::numeric, ",
      db.hdmf,
    ),
    rows(
      sql,
      "tax_table",
      "id, tax_set, tax_code, ",
      "x.doc->>'id', x.doc->>'table', x.doc->>'taxCode', ",
      db.tax,
    ),
    rows(sql, "tax_codes", "id, code, ", "x.doc->>'id', x.doc->>'code', ", db.taxCodes),
    rows(
      sql,
      "user_logs",
      "id, at, username, ",
      "x.doc->>'id', x.doc->>'at', x.doc->>'user', ",
      db.logs,
    ),
    rows(sql, "benefits", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.benefits),
    rows(sql, "assignments", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.assignments),
    rows(sql, "trainings", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.trainings),
    rows(sql, "offenses", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.offenses),
    rows(sql, "leave_balances", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.leaveBalances),
    rows(sql, "appraisals", "id, emp_id, ", "x.doc->>'id', x.doc->>'empId', ", db.appraisals),
    sql.query(`INSERT INTO desk_session (id, session, flash) VALUES (1, $1::jsonb, $2::jsonb)`, [
      JSON.stringify(db.session),
      JSON.stringify(db.flash),
    ]),
  ]);
}
