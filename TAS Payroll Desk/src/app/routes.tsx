import type { ComponentType } from "react";
import type { User } from "../data/types";
import { canRunPayroll, canSee, clientScope } from "../data/rules";
import { Home } from "../screens/Home";
import { ChangePassword } from "../screens/ChangePassword";
import { CompanyScreen } from "../screens/Company";
import { Clients } from "../screens/Clients";
import { Employees } from "../screens/Employees";
import { HumanResources } from "../screens/HumanResources";
import { Shifts } from "../screens/Shifts";
import { Holidays } from "../screens/Holidays";
import { LeaveOvertime } from "../screens/LeaveOvertime";
import { CutoffScreen } from "../screens/Cutoff";
import { ImportDtr } from "../screens/ImportDtr";
import { StaffTime } from "../screens/StaffTime";
import { DriverTime } from "../screens/DriverTime";
import { Deductions } from "../screens/Deductions";
import { PayrollEntry } from "../screens/PayrollEntry";
import { PayslipScreen } from "../screens/Payslip";
import { PayrollHistory } from "../screens/PayrollHistory";
import { Commissions } from "../screens/Commissions";
import { Thirteenth } from "../screens/Thirteenth";
import { Reports } from "../screens/Reports";
import { Maintenance } from "../screens/Maintenance";
import { Users } from "../screens/Users";

export const NAV_GROUPS = [
  "Work",
  "Payroll",
  "Time keeping",
  "People",
  "Reports",
  "Administration",
] as const;

export interface RouteDef {
  path: string;
  href?: string;
  label: string;
  group: (typeof NAV_GROUPS)[number] | null;
  allowed: (u: User) => boolean;
  element: ComponentType<{ param?: string }>;
}

const always = () => true;

export const ROUTES: RouteDef[] = [
  { path: "home", label: "Work queue", group: "Work", allowed: always, element: Home },
  {
    path: "cutoff",
    label: "Cutoff",
    group: "Payroll",
    allowed: (u) => canSee(u, "payroll") || canSee(u, "timeKeeping"),
    element: CutoffScreen,
  },
  {
    path: "payroll-entry",
    label: "Payroll entry",
    group: "Payroll",
    allowed: canRunPayroll,
    element: PayrollEntry,
  },
  {
    path: "payslip",
    label: "Payslip",
    group: "Payroll",
    allowed: (u) => canSee(u, "payroll"),
    element: PayslipScreen,
  },
  {
    path: "payroll-history",
    label: "Payroll history",
    group: "Payroll",
    allowed: (u) => canSee(u, "payroll"),
    element: PayrollHistory,
  },
  {
    path: "commissions",
    label: "Driver commissions",
    group: "Payroll",
    allowed: (u) => canSee(u, "payroll") && clientScope(u) !== "staff",
    element: Commissions,
  },
  {
    path: "thirteenth",
    label: "13th-month payslip",
    group: "Payroll",
    allowed: (u) => canSee(u, "payroll"),
    element: Thirteenth,
  },
  {
    path: "import-dtr",
    label: "Import time records",
    group: "Time keeping",
    allowed: (u) => canSee(u, "timeKeeping"),
    element: ImportDtr,
  },
  {
    path: "time-staff",
    label: "Staff time keeping",
    group: "Time keeping",
    allowed: (u) => canSee(u, "timeKeeping") && clientScope(u) !== "fms",
    element: StaffTime,
  },
  {
    path: "time-fms",
    label: "Driver time keeping",
    group: "Time keeping",
    allowed: (u) => canSee(u, "timeKeeping") && clientScope(u) !== "staff",
    element: DriverTime,
  },
  {
    path: "leave",
    label: "Leave and overtime",
    group: "Time keeping",
    allowed: (u) => canSee(u, "timeKeeping"),
    element: LeaveOvertime,
  },
  {
    path: "holidays",
    label: "Holidays",
    group: "Time keeping",
    allowed: (u) => canSee(u, "timeKeeping"),
    element: Holidays,
  },
  {
    path: "deductions",
    label: "Deductions",
    group: "Time keeping",
    allowed: (u) => canSee(u, "timeKeeping"),
    element: Deductions,
  },
  {
    path: "employees",
    label: "Employees",
    group: "People",
    allowed: (u) => canSee(u, "employeeMaster"),
    element: Employees,
  },
  {
    path: "hr",
    label: "Human resources",
    group: "People",
    allowed: (u) => canSee(u, "employeeInfo"),
    element: HumanResources,
  },
  {
    path: "clients",
    label: "Clients and rates",
    group: "People",
    allowed: (u) => canSee(u, "clientMaster"),
    element: Clients,
  },
  {
    path: "reports",
    label: "Reports",
    group: "Reports",
    allowed: (u) => canSee(u, "reports"),
    element: Reports,
  },
  {
    path: "company",
    label: "Company",
    group: "Administration",
    allowed: (u) => canSee(u, "administrative"),
    element: CompanyScreen,
  },
  {
    path: "shifts",
    label: "Shift schedule",
    group: "Administration",
    allowed: (u) => canSee(u, "administrative"),
    element: Shifts,
  },
  {
    path: "reports",
    href: "reports/user-history",
    label: "User history",
    group: "Administration",
    allowed: (u) => canSee(u, "administrative"),
    element: Reports,
  },
  {
    path: "users",
    label: "Users",
    group: "Administration",
    allowed: (u) => canSee(u, "userAccounts"),
    element: Users,
  },
  {
    path: "maintenance",
    label: "Maintenance tables",
    group: "Administration",
    allowed: (u) => canSee(u, "maintenance"),
    element: Maintenance,
  },
  {
    path: "change-password",
    label: "Change password",
    group: null,
    allowed: always,
    element: ChangePassword,
  },
];

export const routeAllowed = (r: RouteDef, u: User) => r.allowed(u);

export function resolveRoute(path: string, param: string | undefined, u: User) {
  const candidates = ROUTES.filter((r) => r.path === path);
  const match =
    candidates.find((r) => r.href === `${path}/${param}` && r.allowed(u)) ??
    candidates.find((r) => r.allowed(u));
  return { match, exists: candidates.length > 0 };
}
