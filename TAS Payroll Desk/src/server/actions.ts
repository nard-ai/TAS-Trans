import type {
  Client,
  CommissionPayment,
  Company,
  Cutoff,
  Db,
  Deduction,
  Employee,
  PayInputs,
  PayrollRecord,
  Punch,
  User,
} from "../data/types";
import { getDb, mutate } from "./desk";
import { fmtRange } from "../ui/format";
import {
  canOpenClient,
  cutoffHalf,
  cutoffKey,
  fullName,
  isExcluded,
  isFmsClient,
  MONTHS,
} from "../data/rules";
import {
  applyFmsLeave,
  applyLeaveAndOvertime,
  attendanceTotals,
  commissionContributions,
  computePayroll,
  evaluateDay,
  fmsBatch,
  r2,
  shiftContext,
  staffBatch,
  uid,
} from "./engine";

export class ApiError extends Error {
  field?: string;
  constructor(message: string, field?: string) {
    super(message);
    this.field = field;
  }
}

const wait = (ms = 280) => new Promise((r) => setTimeout(r, ms));
const now = () => new Date().toISOString().slice(0, 19);

function currentUser(d: Db): User {
  const u = d.users.find((x) => x.id === d.session?.userId);
  if (!u) throw new ApiError("Your session has ended. Log in again.");
  return u;
}
function log(d: Db, action: string, record: string) {
  const user = d.users.find((x) => x.id === d.session?.userId)?.username ?? "system";
  d.logs.unshift({ id: uid(), at: now(), user, action, record });
}
function requireCutoff(d: Db): { cutoff: Cutoff; client: Client; key: string } {
  const cutoff = d.session?.cutoff;
  if (!cutoff)
    throw new ApiError(
      "No cutoff selected. Choose a client and period on the Cutoff screen first.",
    );
  const client = d.clients.find((c) => c.id === cutoff.clientId)!;
  return { cutoff, client, key: cutoffKey(cutoff.end) };
}
export const cutoffRecord = (d: Db, c: Cutoff | null | undefined) =>
  c ? d.cutoffs.find((r) => r.clientId === c.clientId && r.key === cutoffKey(c.end)) : undefined;
export const isBilled = (d: Db, c: Cutoff | null | undefined) => !!cutoffRecord(d, c)?.billed;
const empLabel = (e: Employee) => `${e.empNo} ${fullName(e)}`;

export type LoginResult =
  | { status: "ok" }
  | { status: "failed"; failures: number }
  | { status: "ended" }
  | { status: "active" }
  | { status: "nodb" };
let failures = 0;

export async function login(
  agencyCode: string,
  username: string,
  password: string,
): Promise<LoginResult> {
  await wait(420);
  const d = getDb();
  const agency = d.agencies.find((a) => a.code === agencyCode);
  if (!agency || !agency.database) return { status: "nodb" };
  const u = d.users.find((x) => x.username.toLowerCase() === username.trim().toLowerCase());
  if (!u || u.password !== password) {
    failures += 1;
    return failures > 3 ? { status: "ended" } : { status: "failed", failures };
  }
  if (u.activeLogin) return { status: "active" };
  failures = 0;
  await mutate((db) => {
    const user = db.users.find((x) => x.id === u.id)!;
    user.activeLogin = true;
    user.lastLoginAt = now();
    db.session = { userId: u.id, agencyCode, loginAt: now(), cutoff: defaultCutoff(db, user) };
    log(db, "Log in", agencyCode);
  });
  return { status: "ok" };
}

function defaultCutoff(db: Db, u: User): Cutoff | null {
  const open = db.cutoffs
    .filter((c) => !c.billed)
    .find((c) => {
      const client = db.clients.find((x) => x.id === c.clientId);
      return client && client.active && canOpenClient(u, client);
    });
  return open ? { clientId: open.clientId, start: open.start, end: open.end } : null;
}

export async function deactivateUser(agencyCode: string, username: string, password: string) {
  await wait(380);
  const d = getDb();
  const u = d.users.find((x) => x.username.toLowerCase() === username.trim().toLowerCase());
  if (!u || u.password !== password) throw new ApiError("Login failed.");
  await mutate((db) => {
    db.users.find((x) => x.id === u.id)!.activeLogin = false;
    db.logs.unshift({
      id: uid(),
      at: now(),
      user: u.username,
      action: "Deactivate user",
      record: `Cleared login lock from ${agencyCode}`,
    });
  });
  return login(agencyCode, username, password);
}

export async function logout() {
  await wait(200);
  await mutate((db) => {
    const u = db.users.find((x) => x.id === db.session?.userId);
    if (u) {
      u.activeLogin = false;
      u.lastLogoutAt = now();
    }
    log(db, "Log out", db.session?.agencyCode ?? "");
    db.session = null;
  });
}

export async function changePassword(current: string, next: string, confirm: string) {
  await wait();
  const d = getDb();
  const u = currentUser(d);
  if (u.password !== current) throw new ApiError("Current password is wrong.", "current");
  if (next.length < 6) throw new ApiError("New password needs at least 6 characters.", "next");
  if (next !== confirm) throw new ApiError("Confirm new password does not match.", "confirm");
  await mutate((db) => {
    db.users.find((x) => x.id === u.id)!.password = next;
    log(db, "Change password", u.username);
  });
}

export async function setFlash(flash: Db["flash"]) {
  await mutate((db) => {
    db.flash = flash;
  });
}

export async function saveCompany(c: Company) {
  await wait();
  if (!c.name.trim()) throw new ApiError("Company name is required.", "name");
  await mutate((db) => {
    db.company = c;
    log(db, "Save company", c.name);
  });
}

export async function saveClient(c: Client) {
  await wait();
  const d = getDb();
  const u = currentUser(d);
  if (!c.code.trim()) throw new ApiError("Client code is required.", "code");
  if (!c.name.trim()) throw new ApiError("Client name is required.", "name");
  if (d.clients.some((x) => x.code === c.code && x.id !== c.id))
    throw new ApiError(`Client code ${c.code} is already used.`, "code");
  if (!canOpenClient(u, c))
    throw new ApiError(
      isFmsClient(c)
        ? "Your account type cannot open FMS clients."
        : "Your account type cannot open staff clients.",
      "code2",
    );
  await mutate((db) => {
    const i = db.clients.findIndex((x) => x.id === c.id);
    if (i >= 0) db.clients[i] = c;
    else db.clients.push(c);
    log(db, i >= 0 ? "Edit client" : "New client", `${c.code} ${c.name}`);
  });
}

export async function saveEmployee(e: Employee) {
  await wait();
  const d = getDb();
  if (!e.empNo.trim()) throw new ApiError("Employee number is required.", "empNo");
  if (!e.lastName.trim()) throw new ApiError("Last name is required.", "lastName");
  if (!e.firstName.trim()) throw new ApiError("First name is required.", "firstName");
  if (!e.clientId) throw new ApiError("Client is required.", "clientId");
  if (d.employees.some((x) => x.empNo === e.empNo && x.id !== e.id))
    throw new ApiError(`Employee number ${e.empNo} is already used.`, "empNo");
  if (e.dtrId && d.employees.some((x) => x.dtrId === e.dtrId && x.id !== e.id))
    throw new ApiError(`DTR id ${e.dtrId} belongs to another employee.`, "dtrId");
  await mutate((db) => {
    const i = db.employees.findIndex((x) => x.id === e.id);
    if (i >= 0) db.employees[i] = e;
    else db.employees.push(e);
    log(db, i >= 0 ? "Edit employee" : "New employee", empLabel(e));
  });
}

export async function copyEmployee(id: string, empNo: string) {
  await wait();
  const d = getDb();
  const src = d.employees.find((x) => x.id === id);
  if (!src) throw new ApiError("Save the employee before copying.");
  if (!empNo.trim()) throw new ApiError("New employee number is required.", "empNo");
  if (d.employees.some((x) => x.empNo === empNo.trim()))
    throw new ApiError(`Employee number ${empNo} is already used.`, "empNo");
  const copy: Employee = {
    ...src,
    id: `e${uid()}`,
    empNo: empNo.trim(),
    dtrId: `D${empNo.trim()}`,
    payrollStatus: "A",
    active: true,
  };
  await mutate((db) => {
    db.employees.push(copy);
    log(db, "Copy employee", `${src.empNo} to ${copy.empNo}`);
  });
  return copy;
}

type RowTable =
  | "shifts"
  | "holidays"
  | "banks"
  | "branches"
  | "departments"
  | "sss"
  | "philhealth"
  | "hdmf"
  | "tax"
  | "taxCodes"
  | "benefits"
  | "assignments"
  | "trainings"
  | "offenses"
  | "leaveBalances"
  | "appraisals"
  | "users";

export async function saveRow<T extends { id: string }>(table: RowTable, row: T, label: string) {
  await wait();
  await mutate((db) => {
    const list = db[table] as unknown as T[];
    const i = list.findIndex((x) => x.id === row.id);
    if (i >= 0) list[i] = row;
    else list.push(row);
    log(db, `${i >= 0 ? "Edit" : "New"} ${label}`, describe(row));
  });
}
export async function deleteRow(table: RowTable, id: string, label: string) {
  await wait();
  await mutate((db) => {
    const list = db[table] as unknown as { id: string }[];
    const i = list.findIndex((x) => x.id === id);
    if (i >= 0) {
      log(db, `Delete ${label}`, describe(list[i]));
      list.splice(i, 1);
    }
  });
}
function describe(row: object) {
  const r = row as Record<string, unknown>;
  return String(r.code ?? r.name ?? r.title ?? r.benefit ?? r.username ?? r.date ?? r.id);
}

export async function deactivateLoginLock(userId: string) {
  await wait();
  await mutate((db) => {
    const u = db.users.find((x) => x.id === userId);
    if (u) {
      u.activeLogin = false;
      log(db, "Deactivate login lock", u.username);
    }
  });
}

export async function selectCutoff(c: Cutoff) {
  await wait(200);
  if (!c.clientId) throw new ApiError("Client is required.", "clientId");
  if (!c.start) throw new ApiError("Start date is required.", "start");
  if (!c.end) throw new ApiError("End date is required.", "end");
  if (c.end < c.start) throw new ApiError("End date is before the start date.", "end");
  if (c.start.slice(0, 7) !== c.end.slice(0, 7))
    throw new ApiError("Start and end must be in the same month.", "end");
  await mutate((db) => {
    if (db.session) db.session.cutoff = c;
    if (!db.cutoffs.some((r) => r.clientId === c.clientId && r.key === cutoffKey(c.end)))
      db.cutoffs.push({
        ...c,
        key: cutoffKey(c.end),
        billed: false,
        billedAt: null,
        billedBy: null,
      });
  });
}

export type PreviewPunch = Omit<Punch, "id"> & { problem: string | null; employee: string | null };

export async function previewPunches(rows: Omit<Punch, "id">[]): Promise<PreviewPunch[]> {
  await wait(350);
  const d = getDb();
  return rows.map((r) => {
    const e = d.employees.find((x) => x.dtrId === r.dtrId);
    let problem: string | null = null;
    if (!e) problem = "DTR id is not on any employee";
    else if (!/^\d{4}-\d{2}-\d{2}$/.test(r.date)) problem = "Date is not valid";
    else if (!/^\d{2}:\d{2}$/.test(r.time)) problem = "Time is not valid";
    else if (r.dir !== "In" && r.dir !== "Out") problem = "Must be In or Out";
    else if (
      d.punches.some(
        (p) => p.dtrId === r.dtrId && p.date === r.date && p.time === r.time && p.dir === r.dir,
      )
    )
      problem = "Already imported";
    return { ...r, problem, employee: e ? empLabel(e) : null };
  });
}

export async function importPunches(rows: PreviewPunch[]) {
  await wait(400);
  const ok = rows.filter((r) => !r.problem);
  if (!ok.length)
    throw new ApiError("No rows can be imported. Every row in the preview has a problem.");
  await mutate((db) => {
    for (const r of ok)
      db.punches.push({
        id: uid(),
        dtrId: r.dtrId,
        date: r.date,
        time: r.time,
        dir: r.dir,
        source: r.source,
      });
    log(db, "Import punches", `${ok.length} rows`);
  });
  return { imported: ok.length, skipped: rows.length - ok.length };
}

export async function replaceDtrId(oldId: string, newId: string) {
  await wait(400);
  const d = getDb();
  const u = currentUser(d);
  if (!oldId.trim()) throw new ApiError("Old DTR id is required.", "old");
  if (!newId.trim()) throw new ApiError("New DTR id is required.", "new");
  if (oldId === newId) throw new ApiError("New DTR id is the same as the old one.", "new");
  const count = d.punches.filter((p) => p.dtrId === oldId).length;
  if (!count) throw new ApiError(`No punches use DTR id ${oldId}.`, "old");
  const owner = d.employees.find((e) => e.dtrId === newId);
  if (!owner)
    throw new ApiError(
      `New DTR id ${newId} is not on any employee, so it cannot be saved. The old rows were kept.`,
      "new",
    );
  await mutate((db) => {
    db.punches.forEach((p) => {
      if (p.dtrId === oldId) p.dtrId = newId;
    });
    log(db, "Replace DTR id", `${oldId} to ${newId} (${count} rows) by ${u.username}`);
  });
  return { count, owner: empLabel(owner) };
}

function empFor(d: Db, empId: string) {
  const e = d.employees.find((x) => x.id === empId);
  if (!e) throw new ApiError("Choose an employee first.");
  return e;
}

export async function computeStaff(empId: string) {
  await wait(500);
  const d = getDb();
  const { cutoff, client, key } = requireCutoff(d);
  if (isFmsClient(client))
    throw new ApiError("This is an FMS client. Use driver and conductor time keeping.");
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = staffBatch(d, e, cutoff);
  await mutate((db) => {
    delete db.staffFinal[`${e.id}|${key}`];
    db.staffFinal[`${e.id}|${key}`] = rows;
    log(db, "Compute time records", `${empLabel(e)}, ${key}`);
  });
}

export async function finalizeStaff(empId: string) {
  await wait(380);
  const d = getDb();
  const { cutoff, key } = requireCutoff(d);
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = d.staffFinal[`${e.id}|${key}`];
  if (!rows) throw new ApiError("Run Compute before finalizing.");
  await mutate((db) => {
    db.payrollAttendance[`${e.id}|${key}`] = {
      kind: "staff",
      staff: rows.filter((r) => r.remark !== "Invalid Day"),
      at: now(),
    };
    log(db, "Finalize for payroll", `${empLabel(e)}, ${key}`);
  });
}

export async function saveManualTime(empId: string, date: string, timeIn: string, timeOut: string) {
  await wait();
  const d = getDb();
  const { cutoff, key } = requireCutoff(d);
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = d.staffFinal[`${e.id}|${key}`];
  if (!rows) throw new ApiError("Run Compute before editing a punch pair.");
  if (!/^\d{2}:\d{2}$/.test(timeIn)) throw new ApiError("Time in is not a valid time.", "in");
  if (!/^\d{2}:\d{2}$/.test(timeOut)) throw new ApiError("Time out is not a valid time.", "out");
  const ctx = shiftContext(d, e);
  await mutate((db) => {
    db.staffFinal[`${e.id}|${key}`] = rows.map((r) => {
      if (r.date !== date) return r;
      const next = { ...r, timeIn, timeOut, manual: true };
      evaluateDay(next, ctx);
      return next;
    });
    log(db, "Save manual time", `${empLabel(e)}, ${date} ${timeIn}-${timeOut}`);
  });
}

export async function refreshStaffLeave(empId: string) {
  await wait(320);
  const d = getDb();
  const { cutoff, key } = requireCutoff(d);
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = d.staffFinal[`${e.id}|${key}`];
  if (!rows) throw new ApiError("Nothing to refresh. Run Compute first.");
  await mutate((db) => {
    db.staffFinal[`${e.id}|${key}`] = applyLeaveAndOvertime(db, e, rows);
    log(db, "Refresh from leave and overtime", `${empLabel(e)}, ${key}`);
  });
}

export async function computeFms(empId: string) {
  await wait(500);
  const d = getDb();
  const { cutoff, client, key } = requireCutoff(d);
  if (!isFmsClient(client)) throw new ApiError("This is a staff client. Use staff time keeping.");
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = fmsBatch(d, e, cutoff);
  await mutate((db) => {
    db.fmsFinal[`${e.id}|${key}`] = rows;
    log(db, "Compute driver time records", `${empLabel(e)}, ${key}`);
  });
}

export async function finalizeFms(empId: string) {
  await wait(380);
  const d = getDb();
  const { cutoff, key } = requireCutoff(d);
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = d.fmsFinal[`${e.id}|${key}`];
  if (!rows) throw new ApiError("Run Compute before finalizing.");
  await mutate((db) => {
    db.payrollAttendance[`${e.id}|${key}`] = { kind: "fms", fms: rows, at: now() };
    log(db, "Finalize for payroll", `${empLabel(e)}, ${key}`);
  });
}

export async function refreshFmsLeave(empId: string) {
  await wait(320);
  const d = getDb();
  const { cutoff, key } = requireCutoff(d);
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed and read-only.");
  const e = empFor(d, empId);
  const rows = d.fmsFinal[`${e.id}|${key}`];
  if (!rows) throw new ApiError("Nothing to refresh. Run Compute first.");
  await mutate((db) => {
    db.fmsFinal[`${e.id}|${key}`] = applyFmsLeave(db, e, rows);
    log(db, "Refresh leave and overtime", `${empLabel(e)}, ${key}`);
  });
}

export async function saveLeave(l: {
  id?: string;
  empId: string;
  date: string;
  letter: string;
  hours: number;
}) {
  await wait();
  if (!l.empId) throw new ApiError("Choose an employee first.");
  if (!l.date) throw new ApiError("Date is required.", "date");
  if (!l.letter) throw new ApiError("Leave letter is required.", "letter");
  if (!(l.hours > 0) || l.hours > 24)
    throw new ApiError("Hours must be between 0 and 24.", "hours");
  await mutate((db) => {
    const existing = db.leaves.find((x) => x.empId === l.empId && x.date === l.date);
    if (existing) Object.assign(existing, { letter: l.letter, hours: l.hours });
    else
      db.leaves.push({
        id: uid(),
        empId: l.empId,
        date: l.date,
        letter: l.letter as never,
        hours: l.hours,
      });
    log(db, "Save leave", `${l.empId} ${l.date} ${l.letter} ${l.hours}h`);
  });
}
export async function deleteLeave(id: string) {
  await wait();
  await mutate((db) => {
    db.leaves = db.leaves.filter((x) => x.id !== id);
    log(db, "Delete leave", id);
  });
}
export async function saveOvertime(o: { empId: string; date: string; hours: number }) {
  await wait();
  if (!o.empId) throw new ApiError("Choose an employee first.");
  if (!o.date) throw new ApiError("Overtime date is required.", "otDate");
  if (!(o.hours >= 0) || o.hours > 16)
    throw new ApiError("Overtime hours must be between 0 and 16.", "otHours");
  await mutate((db) => {
    const existing = db.overtimes.find((x) => x.empId === o.empId && x.date === o.date);
    if (existing) existing.hours = o.hours;
    else db.overtimes.push({ id: uid(), ...o });
    log(db, "Save overtime", `${o.empId} ${o.date} ${o.hours}h`);
  });
}

export async function saveDeduction(row: Deduction) {
  await wait();
  const d = getDb();
  const existing = d.deductions.find((x) => x.id === row.id);
  if (existing?.locked) throw new ApiError("This deduction is locked and cannot be changed.");
  if (!row.code) throw new ApiError("Deduction code is required.", "code");
  if (row.amount < 0) throw new ApiError("Amount this cutoff cannot be negative.", "amount");
  if (row.amount > row.balance)
    throw new ApiError("Amount this cutoff is more than the balance.", "amount");
  await mutate((db) => {
    const i = db.deductions.findIndex((x) => x.id === row.id);
    if (i >= 0) db.deductions[i] = row;
    else db.deductions.push(row);
    log(db, i >= 0 ? "Edit deduction" : "New deduction", `${row.empId} code ${row.code}`);
  });
}
export async function deleteDeduction(id: string) {
  await wait();
  const d = getDb();
  if (d.deductions.find((x) => x.id === id)?.locked)
    throw new ApiError("This deduction is locked and cannot be deleted.");
  await mutate((db) => {
    db.deductions = db.deductions.filter((x) => x.id !== id);
    log(db, "Delete deduction", id);
  });
}

function payrollContext(empId: string) {
  const d = getDb();
  const u = currentUser(d);
  const { cutoff, client, key } = requireCutoff(d);
  const e = empFor(d, empId);
  if (isExcluded(e))
    throw new ApiError(`Payroll status ${e.payrollStatus} is excluded from payroll.`);
  return { d, u, cutoff, client, key, e };
}

export async function loadAttendance(empId: string, inputs: PayInputs) {
  await wait(420);
  const { d, u, cutoff, client, key, e } = payrollContext(empId);
  const att = attendanceTotals(d, e, client, key);
  if (!att)
    throw new ApiError(
      `No DTR records processed for ${fmtRange(cutoff.start, cutoff.end)}. Compute and finalize time keeping for this employee first.`,
    );
  return computePayroll(d, e, client, cutoff, att, inputs, u.username);
}

export async function recompute(empId: string, inputs: PayInputs) {
  return loadAttendance(empId, inputs);
}

export async function savePayroll(empId: string, inputs: PayInputs): Promise<PayrollRecord> {
  await wait(460);
  const { d, u, cutoff, client, key, e } = payrollContext(empId);
  if (isBilled(d, cutoff))
    throw new ApiError("This cutoff is already billed. Saved payroll is read-only.");
  const att = attendanceTotals(d, e, client, key);
  if (!att)
    throw new ApiError(`No DTR records processed for ${fmtRange(cutoff.start, cutoff.end)}.`);
  const { record, problems } = computePayroll(d, e, client, cutoff, att, inputs, u.username);
  if (problems.length) throw new ApiError(problems[0], "taxCode");
  await mutate((db) => {
    db.openPayroll = db.openPayroll.filter(
      (r) => !(r.empId === e.id && r.key === key && r.clientId === client.id),
    );
    db.openPayroll.push(record);
    log(db, "Save payroll", `${empLabel(e)}, cutoff ${key}`);
  });
  return record;
}

export async function moveToMaster() {
  await wait(700);
  const d = getDb();
  const u = currentUser(d);
  const { cutoff, client, key } = requireCutoff(d);
  if (isBilled(d, cutoff)) throw new ApiError("This cutoff is already billed.");
  const rows = d.openPayroll.filter((r) => r.clientId === client.id && r.key === key);
  if (!rows.length)
    throw new ApiError("There is no saved payroll for this client and cutoff to move.");
  await mutate((db) => {
    const at = now();
    db.history.push(...rows.map((r) => ({ ...r, movedAt: at })));
    db.openPayroll = db.openPayroll.filter((r) => !(r.clientId === client.id && r.key === key));
    for (const k of Object.keys(db.payrollAttendance))
      if (
        k.endsWith(`|${key}`) &&
        db.employees.find((e) => e.id === k.split("|")[0])?.clientId === client.id
      )
        delete db.payrollAttendance[k];
    const rec = db.cutoffs.find((r) => r.clientId === client.id && r.key === key);
    if (rec) Object.assign(rec, { billed: true, billedAt: at, billedBy: u.username });
    else db.cutoffs.push({ ...cutoff, key, billed: true, billedAt: at, billedBy: u.username });
    log(db, "Move to master", `${client.name}, cutoff ${key}`);
  });
  return { count: rows.length };
}

export async function loadTrips(empId: string, year: number, month: number, dbName: string) {
  await wait(450);
  if (!empId) throw new ApiError("Choose an employee first.", "empId");
  if (!dbName.trim()) throw new ApiError("FMS database name is required.", "dbName");
  const d = getDb();
  if (!d.trips.some((t) => t.db === dbName.trim()))
    throw new ApiError(`FMS database ${dbName} was not found.`, "dbName");
  const prefix = `${year}-${String(month).padStart(2, "0")}`;
  return d.trips
    .filter((t) => t.db === dbName.trim() && t.empId === empId && t.date.startsWith(prefix))
    .sort((a, b) => a.date.localeCompare(b.date));
}

export async function saveCommissionPayment(p: {
  empId: string;
  year: number;
  month: number;
  commission: number;
  suretyFund: number;
  ceiling: number;
  beginningBalance: number;
  taxAmount: number;
  taxCode: string;
}): Promise<CommissionPayment> {
  await wait(520);
  const d = getDb();
  const u = currentUser(d);
  const e = empFor(d, p.empId);
  if (!p.taxCode) throw new ApiError("Tax code is required.", "taxCode");
  if (!(p.ceiling > 0)) throw new ApiError("Ceiling must be more than zero.", "ceiling");
  const c = p.commission > 0 ? commissionContributions(d, p.commission, p.ceiling) : null;
  const payment: CommissionPayment = {
    id: uid(),
    empId: e.id,
    year: p.year,
    month: p.month,
    commission: r2(p.commission),
    suretyFund: r2(p.suretyFund),
    ceiling: p.ceiling,
    beginningBalance: p.beginningBalance,
    taxAmount: c ? p.taxAmount : 0,
    sssEe: c?.sssEe ?? 0,
    sssEr: c?.sssEr ?? 0,
    phEe: c?.phEe ?? 0,
    phEr: c?.phEr ?? 0,
    hdmfEe: c?.hdmfEe ?? 0,
    hdmfEr: c?.hdmfEr ?? 0,
    ecc: c?.ecc ?? 0,
    contributionsUpdated: !!c,
    savedBy: u.username,
    savedAt: now(),
  };
  await mutate((db) => {
    db.payments = db.payments.filter(
      (x) => !(x.empId === e.id && x.year === p.year && x.month === p.month),
    );
    db.payments.push(payment);
    log(db, "Save commission payment", `${empLabel(e)}, ${MONTHS[p.month - 1]} ${p.year}`);
  });
  return payment;
}

export async function generateThirteenth(empId: string, year: number) {
  await wait(300);
  const d = getDb();
  if (!empId) throw new ApiError("Choose an employee first.", "empId");
  const row = d.thirteenth.find((x) => x.empId === empId && x.year === year);
  if (!row) throw new ApiError(`No stored 13th-month amount for ${year}.`, "year");
  return row;
}

export const halfOf = cutoffHalf;

export const handlers: Record<string, (...args: any[]) => unknown> = {
  login,
  deactivateUser,
  logout,
  changePassword,
  setFlash,
  saveCompany,
  saveClient,
  saveEmployee,
  copyEmployee,
  saveRow,
  deleteRow,
  deactivateLoginLock,
  selectCutoff,
  previewPunches,
  importPunches,
  replaceDtrId,
  computeStaff,
  finalizeStaff,
  saveManualTime,
  refreshStaffLeave,
  computeFms,
  finalizeFms,
  refreshFmsLeave,
  saveLeave,
  deleteLeave,
  saveOvertime,
  saveDeduction,
  deleteDeduction,
  loadAttendance,
  recompute,
  savePayroll,
  moveToMaster,
  loadTrips,
  saveCommissionPayment,
  generateThirteenth,
};
