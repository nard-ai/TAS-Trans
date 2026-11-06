import type { Client, Employee, LeaveLetter, User, Right, Db } from "./types";

/* Display rules shared by every screen. Pay amounts are never computed here. */

export const EXCLUDED_STATUSES = ["T", "W", "X", "Y", "Z"];
export const isExcluded = (e: Employee) => EXCLUDED_STATUSES.includes(e.payrollStatus);
export const isFmsClient = (c: Client | undefined | null) => c?.code2 === "FMS";

export function parseISO(iso: string) {
  const [y, m, d] = iso.split("-").map(Number);
  return { y, m, d };
}
export const pad2 = (n: number) => String(n).padStart(2, "0");
export const isoOf = (y: number, m: number, d: number) => `${y}-${pad2(m)}-${pad2(d)}`;
export const daysInMonth = (y: number, m: number) => new Date(y, m, 0).getDate();

export function weekdayOf(iso: string) {
  const { y, m, d } = parseISO(iso);
  return ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][new Date(y, m - 1, d).getDay()];
}

export function dateRange(start: string, end: string): string[] {
  const out: string[] = [];
  const s = parseISO(start);
  const cur = new Date(s.y, s.m - 1, s.d);
  const e = parseISO(end);
  const last = new Date(e.y, e.m - 1, e.d);
  while (cur <= last && out.length < 62) {
    out.push(isoOf(cur.getFullYear(), cur.getMonth() + 1, cur.getDate()));
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

/** The end day decides the half. */
export const cutoffHalf = (end: string): 1 | 2 => (parseISO(end).d <= 15 ? 1 : 2);
export function cutoffKey(end: string) {
  const { y, m } = parseISO(end);
  return `${y}${pad2(m)}0${cutoffHalf(end)}`;
}
/** End day before the 17th: SSS and PhilHealth show, Pag-IBIG hides. */
export const showsSssPhilhealth = (end: string) => parseISO(end).d < 17;
export const showsPagibig = (end: string) => !showsSssPhilhealth(end);
/** Union dues are zero when the end day is after the 15th. */
export const showsUnionDues = (end: string) => parseISO(end).d <= 15;

export const LEAVE_LETTERS: { letter: LeaveLetter; label: string; reads: string }[] = [
  { letter: "H", label: "Special holiday", reads: "Hours, special-holiday rate" },
  { letter: "L", label: "Legal holiday", reads: "Hours, legal-holiday rate" },
  { letter: "M", label: "Sick leave", reads: "Hours divided by 8, as days" },
  { letter: "S", label: "Sick leave", reads: "Hours divided by 8, as days" },
  { letter: "P", label: "Vacation leave", reads: "Hours divided by 8, as days" },
  { letter: "V", label: "Vacation leave", reads: "Hours divided by 8, as days" },
  { letter: "Y", label: "Vacation leave", reads: "Hours divided by 8, as days" },
  { letter: "R", label: "Rest day", reads: "Hours, subtracted from hours worked" },
  { letter: "D", label: "Day off worked", reads: "Rest-day overtime" },
  { letter: "B", label: "Official business", reads: "Required hours become the leave hours" },
];
export const leaveLabel = (l: string) => LEAVE_LETTERS.find((x) => x.letter === l)?.label ?? "";

export const DEDUCTION_CODES: { code: number; staff: string; fms?: string }[] = [
  { code: 1, staff: "Uniform" },
  { code: 3, staff: "Cash advance" },
  { code: 4, staff: "Planters", fms: "Damage and accident" },
  { code: 5, staff: "SSS loan" },
  { code: 6, staff: "Pag-IBIG loan" },
  { code: 7, staff: "Vale" },
  { code: 8, staff: "Other", fms: "Shortages" },
  { code: 10, staff: "LMC" },
  { code: 11, staff: "LMC" },
  { code: 12, staff: "Surety or calamity", fms: "Surety fund" },
  { code: 16, staff: "HDMF calamity" },
  { code: 17, staff: "SSS calamity" },
];
export function deductionLabel(code: number, fms: boolean) {
  const d = DEDUCTION_CODES.find((x) => x.code === code);
  if (!d) return `Code ${code}`;
  return fms && d.fms ? d.fms : d.staff;
}

export const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

export const fullName = (e: Pick<Employee, "lastName" | "firstName" | "middleName">) =>
  `${e.lastName}, ${e.firstName}${e.middleName ? ` ${e.middleName.charAt(0)}.` : ""}`;

/* Access */

export const canSee = (u: User, r: Right) => u.superuser || u.rights[r];
export const canAdd = (u: User) => u.superuser || u.canAdd;
export const canEdit = (u: User) => u.superuser || u.canEdit;
export const canDelete = (u: User) => u.superuser || u.canDelete;
export const PAYROLL_ACCOUNT_TYPES = [
  "Accounting",
  "Administrator",
  "FMS Payroll",
  "FMS Admin",
  "Staff Admin",
  "Staff Payroll",
];
export const canRunPayroll = (u: User) =>
  PAYROLL_ACCOUNT_TYPES.includes(u.accountType) && canSee(u, "payroll");

export function clientScope(u: User): "fms" | "staff" | "all" {
  const t = u.accountType.toUpperCase();
  if (t.startsWith("FMS")) return "fms";
  if (t.startsWith("STA")) return "staff";
  return "all";
}
export function canOpenClient(u: User, c: Client) {
  const s = clientScope(u);
  if (s === "fms") return isFmsClient(c);
  if (s === "staff") return !isFmsClient(c);
  return true;
}
export const visibleClients = (db: Db, u: User) => db.clients.filter((c) => canOpenClient(u, c));
