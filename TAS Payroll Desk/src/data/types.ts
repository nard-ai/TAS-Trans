export type Right =
  | "payroll"
  | "employeeInfo"
  | "timeKeeping"
  | "clientMaster"
  | "employeeMaster"
  | "userAccounts"
  | "administrative"
  | "reports"
  | "maintenance";

export const RIGHTS: { key: Right; label: string; areas: string }[] = [
  { key: "payroll", label: "Payroll", areas: "Payroll processing, payroll master, payslips" },
  { key: "employeeInfo", label: "Employee information", areas: "Human resources" },
  {
    key: "timeKeeping",
    label: "Time keeping",
    areas: "Import DTR, time keeping, holidays, leave, deductions",
  },
  { key: "clientMaster", label: "Client master", areas: "Client rates" },
  { key: "employeeMaster", label: "Employee master", areas: "Employee payroll profile" },
  { key: "userAccounts", label: "User accounts", areas: "Users" },
  { key: "administrative", label: "Administrative", areas: "Company, shift, user history" },
  { key: "reports", label: "Reports", areas: "HR reports, SSS, PhilHealth, HDMF" },
  {
    key: "maintenance",
    label: "Maintenance",
    areas: "Banks, branches, departments, tax and contribution tables",
  },
];

export const ACCOUNT_TYPES = [
  "Accounting",
  "Administrator",
  "FMS Payroll",
  "FMS Admin",
  "Staff Admin",
  "Staff Payroll",
] as const;
export type AccountType = (typeof ACCOUNT_TYPES)[number];

export interface User {
  id: string;
  username: string;
  password: string;
  name: string;
  accountType: AccountType;
  superuser: boolean;
  rights: Record<Right, boolean>;
  canAdd: boolean;
  canEdit: boolean;
  canDelete: boolean;
  activeLogin: boolean;
  lastLoginAt: string | null;
  lastLogoutAt: string | null;
}

export interface Agency {
  code: string;
  name: string;
  database: string | null;
}

export interface Company {
  name: string;
  address: string;
  tin: string;
  telephone: string;
  sssNo: string;
  philhealthNo: string;
  pagibigNo: string;
}

export interface ClientRates {
  overtime: number;
  nightDiff: number;
  restDay: number;
  specialHoliday: number;
  legalHoliday: number;
  holidayOvertime: number;
  excessOvertime: number;
  ecola: number;
  hazard: number;
  agencyRate: number;
  dailyCharge: number;
}

export interface ClientFlags {
  sss: boolean;
  philhealth: boolean;
  pagibig: boolean;
  thirteenthOnCutoff: boolean;
  displayThirteenth: boolean;
  /** Stored as "ECOLA on rest day"; the same switch selects the 2018 tax table. */
  tax2018: boolean;
}

export interface Client {
  id: string;
  code: string;
  /** "FMS" marks the drivers and conductors client. */
  code2: string;
  name: string;
  active: boolean;
  workingDays: number | null;
  rates: ClientRates;
  flags: ClientFlags;
  taxPercent: boolean;
  taxCeiling: number;
}

export const PAYROLL_STATUSES = ["A", "P", "C", "T", "W", "X", "Y", "Z"] as const;
export type PayrollStatus = (typeof PAYROLL_STATUSES)[number];

export interface Employee {
  id: string;
  empNo: string;
  dtrId: string;
  lastName: string;
  firstName: string;
  middleName: string;
  department: string;
  position: string;
  role: "" | "Driver" | "Conductor";
  taxCode: string;
  civilStatus: string;
  gender: string;
  birthDate: string;
  address: string;
  phone: string;
  tin: string;
  sssNo: string;
  philhealthNo: string;
  pagibigNo: string;
  bank: string;
  accountNo: string;
  basis: "Monthly" | "Hourly";
  monthlyRate: number;
  dailyRate: number;
  hourlyRate: number;
  sss: boolean;
  philhealth: boolean;
  pagibig: boolean;
  philhealthPerDay: number;
  pagibigPerDay: number;
  shift: string;
  restDay: string;
  clientId: string;
  hired: string;
  contractStart: string;
  contractEnd: string;
  payrollStatus: PayrollStatus;
  active: boolean;
}

export interface Shift {
  id: string;
  code: string;
  start: string;
  end: string;
  midday: string;
  workHours: number;
  graceMinutes: number;
  graceOn: boolean;
  noLate: boolean;
  noUndertime: boolean;
  undertimeMinutes: number;
  noInOut: boolean;
  flexible: boolean;
  noOvertime: boolean;
  overtimeStart: string;
  nightDiff: boolean;
}

export interface Holiday {
  id: string;
  date: string;
  name: string;
  letter: "H" | "L";
}

export type LeaveLetter = "H" | "L" | "M" | "S" | "P" | "V" | "Y" | "R" | "D" | "B";

export interface Leave {
  id: string;
  empId: string;
  date: string;
  letter: LeaveLetter;
  hours: number;
}

export interface OvertimeFiling {
  id: string;
  empId: string;
  date: string;
  hours: number;
}

export interface Punch {
  id: string;
  dtrId: string;
  date: string;
  time: string;
  dir: "In" | "Out";
  source: string;
}

export type StaffRemark = "Present" | "Absent" | "No In/Out" | "Invalid Day" | "Day off";

export interface StaffDay {
  date: string;
  weekday: string;
  timeIn: string;
  timeOut: string;
  required: number;
  work: number;
  late: number;
  undertime: number;
  overtime: number;
  nd: number;
  leaveLetter: LeaveLetter | "";
  leaveHours: number;
  remark: StaffRemark;
  manual: boolean;
}

export interface FmsDay {
  day: number;
  date: string | null;
  pairs: { in: string; out: string }[];
  pairHours: number[];
  required: number;
  overtime: number;
  nd: number | null;
  late: number | null;
  leaveLetter: LeaveLetter | "";
  leaveHours: number;
  restDayOt: number;
  remark: StaffRemark;
}

export interface Cutoff {
  clientId: string;
  start: string;
  end: string;
}

export interface CutoffRecord extends Cutoff {
  key: string;
  billed: boolean;
  billedAt: string | null;
  billedBy: string | null;
}

export interface Deduction {
  id: string;
  empId: string;
  code: number;
  balance: number;
  schedule: "Every cutoff" | "First half" | "Second half";
  amount: number;
  locked: boolean;
}

export interface PayLine {
  key: string;
  label: string;
  qty: number | null;
  unit: "hrs" | "days" | "";
  basis: string;
  amount: number;
}

export interface AttendanceTotals {
  kind: "staff" | "fms";
  hoursWorked: number;
  days: number;
  overtime: number;
  nd: number;
  restDay: number;
  special: number;
  legal: number;
  sickDays: number;
  vacationDays: number;
  absentDays: number;
  lateHours: number;
  absenceAmount: number;
  lateAmount: number;
}

export type HourKey = "hoursWorked" | "overtime" | "nd" | "restDay" | "special" | "legal";

export interface PayInputs {
  hours: Partial<Record<HourKey, number>>;
  allowance: number;
  otherIncome: number;
  taxRefund: number;
}

export interface PayrollRecord {
  id: string;
  empId: string;
  clientId: string;
  key: string;
  start: string;
  end: string;
  half: 1 | 2;
  attendance: AttendanceTotals;
  earnings: PayLine[];
  deductions: PayLine[];
  gross: number;
  totalDeductions: number;
  net: number;
  inputs: PayInputs;
  rates: {
    basis: "Monthly" | "Hourly";
    monthly: number;
    daily: number;
    hourly: number;
    taxCode: string;
  };
  processedBy: string;
  savedAt: string;
  movedAt: string | null;
}

export interface Trip {
  id: string;
  db: string;
  empId: string;
  date: string;
  bundle: string;
  commission: number;
  fund: number;
  bus: string;
  route: string;
}

export interface CommissionPayment {
  id: string;
  empId: string;
  year: number;
  month: number;
  commission: number;
  suretyFund: number;
  ceiling: number;
  beginningBalance: number;
  taxAmount: number;
  sssEe: number;
  sssEr: number;
  phEe: number;
  phEr: number;
  hdmfEe: number;
  hdmfEr: number;
  ecc: number;
  contributionsUpdated: boolean;
  savedBy: string;
  savedAt: string;
}

export interface Thirteenth {
  id: string;
  empId: string;
  clientId: string;
  year: number;
  amount: number;
}

export interface Bank {
  id: string;
  code: string;
  name: string;
}
export interface Branch {
  id: string;
  code: string;
  name: string;
  bank: string;
}
export interface Department {
  id: string;
  code: string;
  name: string;
}
export interface SssRow {
  id: string;
  from: number;
  to: number;
  ee: number;
  er: number;
  ec: number;
}
export interface PhRow {
  id: string;
  from: number;
  to: number;
  ee: number;
  er: number;
}
export interface HdmfRow {
  id: string;
  from: number;
  to: number;
  eeRate: number;
  erRate: number;
}
export interface TaxRow {
  id: string;
  table: "Semi-monthly" | "2018";
  taxCode: string;
  floor: number;
  percent: number;
  baseTax: number;
}
export interface TaxCode {
  id: string;
  code: string;
  description: string;
  exemption: number;
}

export interface UserLog {
  id: string;
  at: string;
  user: string;
  action: string;
  record: string;
}

export interface Benefit {
  id: string;
  empId: string;
  benefit: string;
  amount: number;
  start: string;
  end: string;
}
export interface Assignment {
  id: string;
  empId: string;
  client: string;
  position: string;
  start: string;
  end: string;
}
export interface Training {
  id: string;
  empId: string;
  title: string;
  provider: string;
  date: string;
  hours: number;
}
export interface Offense {
  id: string;
  empId: string;
  date: string;
  offense: string;
  action: string;
}
export interface LeaveBalance {
  id: string;
  empId: string;
  year: number;
  type: "Sick leave" | "Vacation leave";
  entitled: number;
  used: number;
}
export interface Appraisal {
  id: string;
  empId: string;
  date: string;
  rating: string;
  rater: string;
  remarks: string;
}

export interface Session {
  userId: string;
  agencyCode: string;
  loginAt: string;
  cutoff: Cutoff | null;
}

export interface Db {
  agencies: Agency[];
  users: User[];
  company: Company | null;
  clients: Client[];
  employees: Employee[];
  shifts: Shift[];
  holidays: Holiday[];
  leaves: Leave[];
  overtimes: OvertimeFiling[];
  punches: Punch[];
  /** Final rows from Compute, keyed `${empId}|${dtrKey}`. */
  staffFinal: Record<string, StaffDay[]>;
  fmsFinal: Record<string, FmsDay[]>;
  /** The set payroll reads, written by Finalize for payroll. */
  payrollAttendance: Record<
    string,
    { kind: "staff" | "fms"; staff?: StaffDay[]; fms?: FmsDay[]; at: string }
  >;
  cutoffs: CutoffRecord[];
  deductions: Deduction[];
  openPayroll: PayrollRecord[];
  history: PayrollRecord[];
  trips: Trip[];
  payments: CommissionPayment[];
  thirteenth: Thirteenth[];
  banks: Bank[];
  branches: Branch[];
  departments: Department[];
  sss: SssRow[];
  philhealth: PhRow[];
  hdmf: HdmfRow[];
  tax: TaxRow[];
  taxCodes: TaxCode[];
  logs: UserLog[];
  benefits: Benefit[];
  assignments: Assignment[];
  trainings: Training[];
  offenses: Offense[];
  leaveBalances: LeaveBalance[];
  appraisals: Appraisal[];
  session: Session | null;
  flash: { tone: "success" | "error" | "info"; title: string } | null;
}
