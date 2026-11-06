import { execFileSync } from "node:child_process";
import { readFileSync, unlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createServer } from "vite";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const r2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
const trim = (v) => String(v ?? "").trim();
const LEAVE = new Set(["P", "V", "Y", "M", "S", "R"]);
const ENDS = [
  "2026-01-10",
  "2026-01-25",
  "2026-02-10",
  "2026-02-25",
  "2026-03-10",
  "2026-03-25",
  "2026-04-10",
  "2026-04-25",
  "2026-05-10",
  "2026-05-25",
  "2026-06-10",
  "2026-06-25",
];

function databaseUrl() {
  const line = readFileSync(join(root, ".env"), "utf8")
    .split(/\r?\n/)
    .find((row) => row.startsWith("DATABASE_URL="));
  if (!line) throw new Error("DATABASE_URL is missing");
  return line.slice("DATABASE_URL=".length).trim().replace(/^"|"$/g, "");
}

function hm(v) {
  const match = trim(v).match(/^(\d{1,2}):(\d{2})/);
  return match ? `${match[1].padStart(2, "0")}:${match[2]}` : "";
}

function punchTime(v) {
  const clock = hm(v);
  if (clock) return clock;
  const digits = trim(v).match(/^(\d{2})(\d{2})/);
  return digits ? `${digits[1]}:${digits[2]}` : "";
}

function nextDay(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  const date = new Date(Date.UTC(y, m - 1, d));
  date.setUTCDate(date.getUTCDate() + 1);
  return date.toISOString().slice(0, 10);
}

function sqlJson(query) {
  const file = join(tmpdir(), `tas-year-${Date.now()}-${Math.random().toString(16).slice(2)}.json`);
  execFileSync(
    "sqlcmd",
    ["-S", "localhost", "-d", "HRAndPayroll_ReverseEngineering", "-E", "-C", "-y", "0", "-o", file, "-Q", `SET NOCOUNT ON; ${query}`],
    { stdio: "inherit" },
  );
  const text = readFileSync(file, "utf8").replace(/^\uFEFF/, "").replace(/\r?\n/g, "").trim();
  unlinkSync(file);
  if (!text || text === "NULL") return [];
  const start = text.indexOf("[");
  const end = text.lastIndexOf("]");
  if (start < 0 || end < start) throw new Error(text.slice(0, 300));
  return JSON.parse(text.slice(start, end + 1));
}

function deductionsFor(start, end, half, known) {
  const rows = sqlJson(`
    SELECT EmployeeID, DeductionCode,
      SUM(DeductionPayment) AS payment,
      SUM(ISNULL(RunningBalance, DeductionPayment)) AS balance
    FROM EmployeeDedSchedules
    WHERE DeductionStatus = 1
      AND CONVERT(varchar(8), DedStartPayment, 112) >= '${start.replaceAll("-", "")}'
      AND CONVERT(varchar(8), DedEndPayment, 112) <= '${end.replaceAll("-", "")}'
    GROUP BY EmployeeID, DeductionCode
    FOR JSON PATH
  `);
  const schedule = half === 1 ? "First half" : "Second half";
  return rows.flatMap((row) => {
    const empId = `e${row.EmployeeID}`;
    if (!known.has(empId)) return [];
    const amount = r2(row.payment);
    if (amount <= 0) return [];
    return [{
      id: `s${row.EmployeeID}-${row.DeductionCode}`,
      empId,
      code: row.DeductionCode,
      amount,
      balance: Math.max(amount, r2(row.balance)),
      schedule,
      locked: false,
    }];
  });
}

const vite = await createServer({
  root,
  configFile: false,
  server: { middlewareMode: true },
  appType: "custom",
  logLevel: "error",
});
const engine = await vite.ssrLoadModule("/src/server/engine.ts");
const rules = await vite.ssrLoadModule("/src/data/rules.ts");
const persist = await vite.ssrLoadModule("/src/server/persist.ts");
const sql = persist.connect(databaseUrl());
const db = await persist.loadDb(sql);
if (!db) throw new Error("The payroll database is empty.");
const processedBy = db.users.find((user) => user.id === db.session?.userId)?.username || "admin";
const known = new Set(db.employees.map((emp) => emp.id));
const byDtr = new Map();
for (const emp of db.employees) if (emp.dtrId && !byDtr.has(emp.dtrId)) byDtr.set(emp.dtrId, emp);
const problems = [];

db.staffFinal = {};
db.fmsFinal = {};
db.payrollAttendance = {};

for (const end of ENDS) {
  const pending = db.cutoffs.filter(
    (cutoff) =>
      cutoff.end === end &&
      !db.history.some((row) => row.clientId === cutoff.clientId && row.key === cutoff.key),
  );
  if (!pending.length) {
    console.log(`${end} already in history`);
    continue;
  }
  const starts = [...new Set(pending.map((cutoff) => cutoff.start))];
  if (starts.length !== 1) throw new Error(`${end} has more than one start date`);
  const start = starts[0];
  const key = pending[0].key;
  const half = rules.cutoffHalf(end);
  const at = new Date().toISOString();
  const punchRows = sqlJson(`
    SELECT LTRIM(RTRIM(DTRID)) AS DTRID,
      CONVERT(varchar(10), DTR_Date, 23) AS punchDate,
      CONVERT(varchar(8), DTR_Time, 108) AS DTR_Time,
      DTR_IOMode, Machine_ID
    FROM DTRRecords
    WHERE DTR_Date >= '${start}' AND DTR_Date < '${nextDay(end)}'
    FOR JSON PATH
  `);
  db.punches = [];
  for (const row of punchRows) {
    const time = punchTime(row.DTR_Time);
    const dtrId = trim(row.DTRID);
    if (!time || !dtrId || !row.punchDate) continue;
    db.punches.push({
      id: `p${db.punches.length + 1}`,
      dtrId,
      date: row.punchDate,
      time,
      dir: Number(row.DTR_IOMode) === 1 ? "Out" : "In",
      source: row.Machine_ID == null ? "" : String(row.Machine_ID),
    });
  }
  const leaveRows = sqlJson(`
    SELECT LTRIM(RTRIM(DTRID)) AS DTRID,
      CONVERT(varchar(10), DTR_Date, 23) AS leaveDate,
      LTRIM(RTRIM(LeaveType)) AS LeaveType,
      LeaveHours
    FROM DTRPayrollRecords
    WHERE DTRKey = '${key}'
      AND LeaveType IN ('P','V','Y','M','S','R')
      AND LeaveHours > 0
    FOR JSON PATH
  `);
  const leaveMap = new Map();
  for (const row of leaveRows) {
    const emp = byDtr.get(trim(row.DTRID));
    const letter = trim(row.LeaveType);
    const hours = r2(row.LeaveHours);
    if (!emp || !LEAVE.has(letter) || hours <= 0 || !row.leaveDate) continue;
    const id = `${emp.id}|${row.leaveDate}`;
    const prev = leaveMap.get(id);
    if (!prev || hours > prev.hours)
      leaveMap.set(id, { id: `lv-${key}-${emp.id}-${row.leaveDate}`, empId: emp.id, date: row.leaveDate, letter, hours });
  }
  db.leaves = [...leaveMap.values()];
  db.deductions = deductionsFor(start, end, half, known);
  let saved = 0;
  let gross = 0;
  let net = 0;
  for (const cutoff of pending) {
    const client = db.clients.find((row) => row.id === cutoff.clientId);
    const people = db.employees.filter((emp) => emp.clientId === cutoff.clientId && !rules.isExcluded(emp));
    const fms = rules.isFmsClient(client);
    for (const emp of people) {
      const rows = fms ? engine.fmsBatch(db, emp, cutoff) : engine.staffBatch(db, emp, cutoff);
      const attendanceKey = `${emp.id}|${cutoff.key}`;
      db.payrollAttendance[attendanceKey] = fms
        ? { kind: "fms", fms: rows, at }
        : { kind: "staff", staff: rows.filter((row) => row.remark !== "Invalid Day"), at };
      const att = engine.attendanceTotals(db, emp, client, cutoff.key);
      delete db.payrollAttendance[attendanceKey];
      if (!att) {
        problems.push(`${emp.empNo} ${cutoff.key} no attendance`);
        continue;
      }
      const result = engine.computePayroll(
        db,
        emp,
        client,
        cutoff,
        att,
        { hours: {}, allowance: 0, otherIncome: 0, taxRefund: 0 },
        processedBy,
      );
      if (result.problems.length) {
        problems.push(`${emp.empNo} ${client.name}: ${result.problems[0]}`);
        continue;
      }
      db.history.push({ ...result.record, movedAt: at });
      saved += 1;
      gross += result.record.gross;
      net += result.record.net;
    }
  }
  db.punches = [];
  db.leaves = [];
  db.staffFinal = {};
  db.fmsFinal = {};
  db.payrollAttendance = {};
  db.logs.push({
    id: `log-${key}-${Date.now()}`,
    at,
    user: processedBy,
    action: "Process cutoff",
    record: `${start} to ${end}, ${saved} payslips`,
  });
  await persist.saveDb(sql, db);
  console.log(`${start} to ${end} ${key} saved=${saved} gross=${r2(gross)} net=${r2(net)} punches=${punchRows.length} leave=${db.logs.at(-1).record}`);
  console.log(`  leaveRows=${leaveRows.length} deductions=${db.deductions.length}`);
}

db.deductions = deductionsFor("2026-06-26", "2026-07-10", 1, known);
db.punches = [];
db.leaves = [];
await persist.saveDb(sql, db);
const keys = [...new Set(db.history.map((row) => row.key))].sort();
console.log(`history=${db.history.length} keys=${keys.join(",")} problems=${problems.length}`);
for (const line of problems) console.log(`FAIL ${line}`);
await vite.close();
