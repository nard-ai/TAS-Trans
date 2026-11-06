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
  Trip,
} from "./types";
import { cutoffKey } from "./rules";
import { replaceDb } from "./store";

export class ApiError extends Error {
  field?: string;
  constructor(message: string, field?: string) {
    super(message);
    this.field = field;
  }
}

export const cutoffRecord = (d: Db, c: Cutoff | null | undefined) =>
  c ? d.cutoffs.find((r) => r.clientId === c.clientId && r.key === cutoffKey(c.end)) : undefined;
export const isBilled = (d: Db, c: Cutoff | null | undefined) => !!cutoffRecord(d, c)?.billed;

export type LoginResult =
  | { status: "ok" }
  | { status: "failed"; failures: number }
  | { status: "ended" }
  | { status: "active" }
  | { status: "nodb" };

export type PreviewPunch = Omit<Punch, "id"> & { problem: string | null; employee: string | null };

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

async function call<T>(name: string, args: unknown[]): Promise<T> {
  const res = await fetch("/api/call", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ name, args }),
  });
  let body: { result?: T; db?: Db; error?: string; field?: string } = {};
  try {
    body = (await res.json()) as typeof body;
  } catch {
    throw new ApiError("The payroll database did not answer.");
  }
  if (!res.ok) throw new ApiError(body.error || "The payroll database did not answer.", body.field);
  if (body.db) replaceDb(body.db);
  return body.result as T;
}

export const login = (agencyCode: string, username: string, password: string) =>
  call<LoginResult>("login", [agencyCode, username, password]);
export const deactivateUser = (agencyCode: string, username: string, password: string) =>
  call<LoginResult>("deactivateUser", [agencyCode, username, password]);
export const logout = () => call<void>("logout", []);
export const changePassword = (current: string, next: string, confirm: string) =>
  call<void>("changePassword", [current, next, confirm]);
export const setFlash = (flash: Db["flash"]) => {
  void call("setFlash", [flash]).catch(() => undefined);
};
export const saveCompany = (c: Company) => call<void>("saveCompany", [c]);
export const saveClient = (c: Client) => call<void>("saveClient", [c]);
export const saveEmployee = (e: Employee) => call<void>("saveEmployee", [e]);
export const copyEmployee = (id: string, empNo: string) =>
  call<Employee>("copyEmployee", [id, empNo]);
export const saveRow = <T extends { id: string }>(table: RowTable, row: T, label: string) =>
  call<void>("saveRow", [table, row, label]);
export const deleteRow = (table: RowTable, id: string, label: string) =>
  call<void>("deleteRow", [table, id, label]);
export const deactivateLoginLock = (userId: string) => call<void>("deactivateLoginLock", [userId]);
export const selectCutoff = (c: Cutoff) => call<void>("selectCutoff", [c]);
export const previewPunches = (rows: Omit<Punch, "id">[]) =>
  call<PreviewPunch[]>("previewPunches", [rows]);
export const importPunches = (rows: PreviewPunch[]) =>
  call<{ imported: number; skipped: number }>("importPunches", [rows]);
export const replaceDtrId = (oldId: string, newId: string) =>
  call<{ count: number; owner: string }>("replaceDtrId", [oldId, newId]);
export const computeStaff = (empId: string) => call<void>("computeStaff", [empId]);
export const finalizeStaff = (empId: string) => call<void>("finalizeStaff", [empId]);
export const saveManualTime = (empId: string, date: string, timeIn: string, timeOut: string) =>
  call<void>("saveManualTime", [empId, date, timeIn, timeOut]);
export const refreshStaffLeave = (empId: string) => call<void>("refreshStaffLeave", [empId]);
export const computeFms = (empId: string) => call<void>("computeFms", [empId]);
export const finalizeFms = (empId: string) => call<void>("finalizeFms", [empId]);
export const refreshFmsLeave = (empId: string) => call<void>("refreshFmsLeave", [empId]);
export const saveLeave = (l: {
  id?: string;
  empId: string;
  date: string;
  letter: string;
  hours: number;
}) => call<void>("saveLeave", [l]);
export const deleteLeave = (id: string) => call<void>("deleteLeave", [id]);
export const saveOvertime = (o: { empId: string; date: string; hours: number }) =>
  call<void>("saveOvertime", [o]);
export const saveDeduction = (row: Deduction) => call<void>("saveDeduction", [row]);
export const deleteDeduction = (id: string) => call<void>("deleteDeduction", [id]);
export const loadAttendance = (empId: string, inputs: PayInputs) =>
  call<{ record: PayrollRecord; problems: string[] }>("loadAttendance", [empId, inputs]);
export const recompute = (empId: string, inputs: PayInputs) =>
  call<{ record: PayrollRecord; problems: string[] }>("recompute", [empId, inputs]);
export const savePayroll = (empId: string, inputs: PayInputs) =>
  call<PayrollRecord>("savePayroll", [empId, inputs]);
export const moveToMaster = () => call<{ count: number }>("moveToMaster", []);
export const loadTrips = (empId: string, year: number, month: number, dbName: string) =>
  call<Trip[]>("loadTrips", [empId, year, month, dbName]);
export const saveCommissionPayment = (p: {
  empId: string;
  year: number;
  month: number;
  commission: number;
  suretyFund: number;
  ceiling: number;
  beginningBalance: number;
  taxAmount: number;
  taxCode: string;
}) => call<CommissionPayment>("saveCommissionPayment", [p]);
export const generateThirteenth = (empId: string, year: number) =>
  call<{ id: string; empId: string; clientId: string; year: number; amount: number }>(
    "generateThirteenth",
    [empId, year],
  );
