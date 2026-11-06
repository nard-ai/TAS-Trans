import path from "node:path";
import { fileURLToPath } from "node:url";
import { createServer } from "vite";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const base = "http://localhost:8443";
const inputs = { hours: {}, allowance: 0, otherIncome: 0, taxRefund: 0 };

const vite = await createServer({
  root,
  configFile: false,
  server: { middlewareMode: true },
  appType: "custom",
  logLevel: "error",
});
const engine = await vite.ssrLoadModule("/src/server/engine.ts");
const rules = await vite.ssrLoadModule("/src/data/rules.ts");

const got = await fetch(`${base}/api/db`).then((r) => r.json());
const db = got.db;
const user = db.users.find((u) => u.id === db.session?.userId);
const processedBy = user?.username || "admin";
const at = new Date().toISOString();
const pending = db.cutoffs.filter(
  (c) =>
    c.end === "2026-07-10" &&
    !db.history.some((h) => h.clientId === c.clientId && h.key === c.key),
);
const failures = [];

for (const cutoff of pending) {
  const client = db.clients.find((c) => c.id === cutoff.clientId);
  const people = db.employees.filter(
    (e) => e.clientId === cutoff.clientId && !rules.isExcluded(e),
  );
  let saved = 0;
  let gross = 0;
  let net = 0;
  for (const emp of people) {
    const rows = engine.staffBatch(db, emp, cutoff);
    db.staffFinal[`${emp.id}|${cutoff.key}`] = rows;
    db.payrollAttendance[`${emp.id}|${cutoff.key}`] = {
      kind: "staff",
      staff: rows.filter((r) => r.remark !== "Invalid Day"),
      at,
    };
    const att = engine.attendanceTotals(db, emp, client, cutoff.key);
    if (!att) {
      failures.push(`${emp.empNo} ${cutoff.key} no attendance`);
      continue;
    }
    const result = engine.computePayroll(db, emp, client, cutoff, att, inputs, processedBy);
    if (result.problems.length) {
      failures.push(`${emp.empNo} ${client.name}: ${result.problems[0]}`);
      continue;
    }
    db.history.push({ ...result.record, movedAt: at });
    saved += 1;
    gross += result.record.gross;
    net += result.record.net;
  }
  for (const key of Object.keys(db.payrollAttendance)) {
    if (!key.endsWith(`|${cutoff.key}`)) continue;
    const emp = db.employees.find((e) => e.id === key.split("|")[0]);
    if (emp?.clientId === cutoff.clientId) delete db.payrollAttendance[key];
  }
  console.log(
    `${client.name} saved=${saved}/${people.length} gross=${engine.r2(gross)} net=${engine.r2(net)}`,
  );
}

const body = JSON.stringify(db);
console.log(`cutoffs=${pending.length} bytes=${body.length} failures=${failures.length}`);
for (const line of failures) console.log(`FAIL ${line}`);
if (body.length > 7_500_000) throw new Error("Payroll save is over the size limit. Nothing was written.");
const put = await fetch(`${base}/api/db`, {
  method: "PUT",
  headers: { "content-type": "application/json" },
  body,
});
if (!put.ok) throw new Error(await put.text());
console.log("saved");
await vite.close();
