import { execFileSync } from "node:child_process";
import { readFileSync, unlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createServer } from "vite";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const base = "http://localhost:8443";
const r2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;

function sqlJson(query) {
  const file = join(tmpdir(), `tas-ded-${Date.now()}.json`);
  execFileSync(
    "sqlcmd",
    ["-S", "localhost", "-d", "HRAndPayroll_ReverseEngineering", "-E", "-C", "-y", "0", "-o", file, "-Q", `SET NOCOUNT ON; ${query}`],
    { stdio: "inherit" },
  );
  const text = readFileSync(file, "utf8").replace(/^\uFEFF/, "").replace(/\r?\n/g, "").trim();
  unlinkSync(file);
  const start = text.indexOf("[");
  const end = text.lastIndexOf("]");
  if (start < 0 || end < start) throw new Error(text.slice(0, 300));
  return JSON.parse(text.slice(start, end + 1));
}

const rows = sqlJson(`
  SELECT EmployeeID, DeductionCode,
    SUM(DeductionPayment) AS payment,
    SUM(ISNULL(RunningBalance, DeductionPayment)) AS balance
  FROM EmployeeDedSchedules
  WHERE DeductionStatus = 1
    AND CONVERT(varchar(8), DedStartPayment, 112) >= '20260626'
    AND CONVERT(varchar(8), DedEndPayment, 112) <= '20260710'
  GROUP BY EmployeeID, DeductionCode
  FOR JSON PATH
`);

const vite = await createServer({
  root,
  configFile: false,
  server: { middlewareMode: true },
  appType: "custom",
  logLevel: "error",
});
const engine = await vite.ssrLoadModule("/src/server/engine.ts");
const got = await fetch(`${base}/api/db`).then((r) => r.json());
const db = got.db;
const known = new Set(db.employees.map((e) => e.id));
const deductions = [];
let skippedPeople = 0;
for (const row of rows) {
  const empId = `e${row.EmployeeID}`;
  if (!known.has(empId)) {
    skippedPeople += 1;
    continue;
  }
  const amount = r2(row.payment);
  if (amount <= 0) continue;
  deductions.push({
    id: `s${row.EmployeeID}-${row.DeductionCode}`,
    empId,
    code: row.DeductionCode,
    amount,
    balance: Math.max(amount, r2(row.balance)),
    schedule: "First half",
    locked: false,
  });
}
db.deductions = deductions;
const withLoan = new Set(deductions.map((d) => d.empId));
const kept = ["sss", "philhealth", "pagibig", "tax", "union", "absence", "late"];
const amountOf = (rec, key) => rec.deductions.find((d) => d.key === key)?.amount ?? 0;
let updated = 0;
let added = 0;
let netBefore = 0;
let netAfter = 0;
const problems = [];

for (let i = 0; i < db.history.length; i++) {
  const old = db.history[i];
  if (old.key !== "20260701" || !withLoan.has(old.empId)) continue;
  const emp = db.employees.find((e) => e.id === old.empId);
  const client = db.clients.find((c) => c.id === old.clientId);
  const cutoff = db.cutoffs.find((c) => c.clientId === old.clientId && c.key === old.key);
  const result = engine.computePayroll(db, emp, client, cutoff, old.attendance, old.inputs, old.processedBy);
  if (result.problems.length) {
    problems.push(`${emp.empNo}: ${result.problems[0]}`);
    continue;
  }
  const next = result.record;
  const drift = next.gross !== old.gross || kept.some((k) => amountOf(next, k) !== amountOf(old, k));
  if (drift) {
    problems.push(`${emp.empNo}: gross or statutory amount changed`);
    continue;
  }
  const loan = r2(next.totalDeductions - old.totalDeductions);
  db.history[i] = { ...next, id: old.id, savedAt: old.savedAt, movedAt: old.movedAt, processedBy: old.processedBy };
  updated += 1;
  added += loan;
  netBefore += old.net;
  netAfter += next.net;
}

console.log(
  `schedules=${rows.length} stored=${deductions.length} unknown=${skippedPeople} payslips=${updated} loans=${r2(added)} net ${r2(netBefore)} -> ${r2(netAfter)} problems=${problems.length}`,
);
for (const line of problems) console.log(`FAIL ${line}`);
const body = JSON.stringify(db);
console.log(`bytes=${body.length}`);
if (problems.length) throw new Error("Payslips were left unchanged.");
if (body.length > 7_500_000) throw new Error("Payroll save is over the size limit. Nothing was written.");
const put = await fetch(`${base}/api/db`, {
  method: "PUT",
  headers: { "content-type": "application/json" },
  body,
});
if (!put.ok) throw new Error(await put.text());
console.log("saved");
await vite.close();
