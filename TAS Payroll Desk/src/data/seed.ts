import type {
  Client,
  ClientFlags,
  ClientRates,
  Db,
  Employee,
  Punch,
  Right,
  SssRow,
  PhRow,
  TaxRow,
  Trip,
  User,
  CutoffRecord,
  Cutoff,
} from "./types";
import { cutoffKey, dateRange, isExcluded, weekdayOf } from "./rules";
import {
  attendanceTotals,
  commissionContributions,
  computePayroll,
  emptyInputs,
  fmsBatch,
  r2,
  rateBook,
  staffBatch,
} from "../server/engine";

function rng(seed: number) {
  let s = seed % 2147483647;
  if (s <= 0) s += 2147483646;
  return () => (s = (s * 16807) % 2147483647) / 2147483647;
}
const t = (min: number) =>
  `${String(Math.floor(min / 60) % 24).padStart(2, "0")}:${String(min % 60).padStart(2, "0")}`;

const allRights = (on: boolean): Record<Right, boolean> => ({
  payroll: on,
  employeeInfo: on,
  timeKeeping: on,
  clientMaster: on,
  employeeMaster: on,
  userAccounts: on,
  administrative: on,
  reports: on,
  maintenance: on,
});

const baseRates: ClientRates = {
  overtime: 1.25,
  nightDiff: 0.1,
  restDay: 1.3,
  specialHoliday: 1.3,
  legalHoliday: 2,
  holidayOvertime: 2.6,
  excessOvertime: 1.69,
  ecola: 0,
  hazard: 0,
  agencyRate: 0,
  dailyCharge: 0,
};
const baseFlags: ClientFlags = {
  sss: true,
  philhealth: true,
  pagibig: true,
  thirteenthOnCutoff: false,
  displayThirteenth: false,
  tax2018: true,
};

function emp(
  p: Partial<Employee> & Pick<Employee, "id" | "empNo" | "lastName" | "firstName" | "clientId">,
): Employee {
  return {
    dtrId: `D${p.empNo}`,
    middleName: "",
    department: "Operations",
    position: "",
    role: "",
    taxCode: "S",
    civilStatus: "Single",
    gender: "Male",
    birthDate: "1990-01-01",
    address: "Quezon City",
    phone: "",
    tin: "",
    sssNo: "",
    philhealthNo: "",
    pagibigNo: "",
    bank: "BDO",
    accountNo: "",
    basis: "Hourly",
    monthlyRate: 0,
    dailyRate: 610,
    hourlyRate: 0,
    sss: true,
    philhealth: true,
    pagibig: true,
    philhealthPerDay: 0,
    pagibigPerDay: 0,
    shift: "DAY8",
    restDay: "Sun",
    hired: "2021-03-01",
    contractStart: "2025-01-01",
    contractEnd: "2025-12-31",
    payrollStatus: "A",
    active: true,
    ...p,
  };
}

export function seed(): Db {
  const users: User[] = [
    {
      id: "u1",
      username: "admin",
      password: "tas123",
      name: "Rowena Castillo",
      accountType: "Administrator",
      superuser: true,
      rights: allRights(true),
      canAdd: true,
      canEdit: true,
      canDelete: true,
      activeLogin: false,
      lastLoginAt: "2025-12-01T08:02:00",
      lastLogoutAt: "2025-12-01T17:31:00",
    },
    {
      id: "u2",
      username: "staffpay",
      password: "tas123",
      name: "Marites Uy",
      accountType: "Staff Payroll",
      superuser: false,
      rights: {
        ...allRights(false),
        payroll: true,
        timeKeeping: true,
        employeeMaster: true,
        reports: true,
      },
      canAdd: true,
      canEdit: true,
      canDelete: false,
      activeLogin: false,
      lastLoginAt: "2025-12-01T07:55:00",
      lastLogoutAt: "2025-12-01T17:05:00",
    },
    {
      id: "u3",
      username: "fmspay",
      password: "tas123",
      name: "Dennis Yap",
      accountType: "FMS Payroll",
      superuser: false,
      rights: { ...allRights(false), payroll: true, timeKeeping: true },
      canAdd: true,
      canEdit: true,
      canDelete: false,
      activeLogin: false,
      lastLoginAt: "2025-11-28T08:10:00",
      lastLogoutAt: "2025-11-28T18:00:00",
    },
    {
      id: "u4",
      username: "jbernardo",
      password: "tas123",
      name: "Joy Bernardo",
      accountType: "Staff Admin",
      superuser: false,
      rights: { ...allRights(true), userAccounts: false, maintenance: false },
      canAdd: true,
      canEdit: true,
      canDelete: true,
      activeLogin: true,
      lastLoginAt: "2025-12-02T07:48:00",
      lastLogoutAt: "2025-12-01T17:20:00",
    },
    {
      id: "u5",
      username: "lpanganiban",
      password: "tas123",
      name: "Leo Panganiban",
      accountType: "Accounting",
      superuser: false,
      rights: { ...allRights(false), payroll: true, reports: true },
      canAdd: false,
      canEdit: true,
      canDelete: false,
      activeLogin: false,
      lastLoginAt: "2025-11-30T09:00:00",
      lastLogoutAt: "2025-11-30T16:00:00",
    },
  ];

  const clients: Client[] = [
    {
      id: "c1",
      code: "HO",
      code2: "",
      name: "TAS Head Office",
      active: true,
      workingDays: 313,
      rates: { ...baseRates, ecola: 0, hazard: 0 },
      flags: { ...baseFlags, thirteenthOnCutoff: true, displayThirteenth: true },
      taxPercent: true,
      taxCeiling: 610,
    },
    {
      id: "c2",
      code: "CBT",
      code2: "",
      name: "Cubao Terminal",
      active: true,
      workingDays: null,
      rates: { ...baseRates, ecola: 15, hazard: 20 },
      flags: { ...baseFlags, tax2018: false },
      taxPercent: false,
      taxCeiling: 0,
    },
    {
      id: "c3",
      code: "FLT",
      code2: "FMS",
      name: "TAS Fleet Drivers and Conductors",
      active: true,
      workingDays: 365,
      rates: { ...baseRates, restDay: 1.3 },
      flags: { ...baseFlags },
      taxPercent: true,
      taxCeiling: 650,
    },
    {
      id: "c4",
      code: "BTG",
      code2: "",
      name: "Batangas Depot",
      active: false,
      workingDays: 261,
      rates: baseRates,
      flags: baseFlags,
      taxPercent: false,
      taxCeiling: 0,
    },
  ];

  const employees: Employee[] = [
    emp({
      id: "e1001",
      empNo: "1001",
      lastName: "Dela Cruz",
      firstName: "Juan",
      middleName: "Pascual",
      clientId: "c1",
      department: "Accounting",
      position: "Accounting clerk",
      dailyRate: 610,
      tin: "412-551-208-000",
      sssNo: "34-5521890-2",
      hired: "2019-06-17",
    }),
    emp({
      id: "e1002",
      empNo: "1002",
      lastName: "Santos",
      firstName: "Maria",
      middleName: "Luna",
      clientId: "c1",
      department: "Human resources",
      position: "HR assistant",
      gender: "Female",
      basis: "Monthly",
      monthlyRate: 22000,
      dailyRate: 0,
      taxCode: "ME1",
      civilStatus: "Married",
      hired: "2020-02-03",
    }),
    emp({
      id: "e1003",
      empNo: "1003",
      lastName: "Reyes",
      firstName: "Antonio",
      middleName: "Bautista",
      clientId: "c1",
      department: "Operations",
      position: "Dispatcher",
      dailyRate: 645,
      hired: "2018-09-10",
    }),
    emp({
      id: "e1004",
      empNo: "1004",
      lastName: "Garcia",
      firstName: "Liza",
      middleName: "Mercado",
      clientId: "c1",
      department: "Accounting",
      position: "Cashier",
      gender: "Female",
      dailyRate: 610,
      taxCode: "ME",
      civilStatus: "Married",
      hired: "2022-01-10",
    }),
    emp({
      id: "e1005",
      empNo: "1005",
      lastName: "Mendoza",
      firstName: "Ramon",
      middleName: "Tolentino",
      clientId: "c1",
      department: "Maintenance",
      position: "Mechanic",
      dailyRate: 700,
      taxCode: "S1",
      hired: "2017-04-24",
    }),
    emp({
      id: "e1006",
      empNo: "1006",
      lastName: "Bautista",
      firstName: "Carmela",
      middleName: "Ramos",
      clientId: "c1",
      department: "Accounting",
      position: "Payroll officer",
      gender: "Female",
      basis: "Monthly",
      monthlyRate: 28500,
      dailyRate: 0,
      hired: "2016-08-01",
      contractEnd: "",
    }),
    emp({
      id: "e1007",
      empNo: "1007",
      lastName: "Villanueva",
      firstName: "Pedro",
      middleName: "Diaz",
      clientId: "c1",
      department: "Security",
      position: "Security guard",
      dailyRate: 610,
      shift: "NIGHT",
      restDay: "Mon",
      philhealthPerDay: 15,
      hired: "2023-05-02",
    }),
    emp({
      id: "e1008",
      empNo: "1008",
      lastName: "Fernandez",
      firstName: "Grace",
      middleName: "Abad",
      clientId: "c1",
      department: "Operations",
      position: "Encoder",
      gender: "Female",
      payrollStatus: "T",
      active: false,
      contractEnd: "2025-09-30",
    }),
    emp({
      id: "e1009",
      empNo: "1009",
      lastName: "Aquino",
      firstName: "Mark",
      middleName: "Javier",
      clientId: "c1",
      department: "Operations",
      position: "Utility",
      dailyRate: 590,
      taxCode: "S",
      hired: "2025-10-01",
      contractStart: "2025-10-01",
      contractEnd: "2026-03-31",
    }),
    emp({
      id: "e1010",
      empNo: "1010",
      lastName: "Ramos",
      firstName: "Teresa",
      middleName: "Cruz",
      clientId: "c1",
      department: "Human resources",
      position: "Admin assistant",
      gender: "Female",
      dailyRate: 620,
      taxCode: "S",
      hired: "2024-07-15",
      contractEnd: "2025-12-15",
    }),
    emp({
      id: "e1011",
      empNo: "1011",
      lastName: "Castillo",
      firstName: "Jose",
      middleName: "Esteban",
      clientId: "c1",
      department: "Operations",
      position: "Driver liaison",
      payrollStatus: "W",
      active: false,
    }),
    emp({
      id: "e2001",
      empNo: "2001",
      lastName: "Torres",
      firstName: "Ana",
      middleName: "Santiago",
      clientId: "c2",
      department: "Operations",
      position: "Ticket teller",
      gender: "Female",
      dailyRate: 645,
      hired: "2021-11-08",
    }),
    emp({
      id: "e2002",
      empNo: "2002",
      lastName: "Navarro",
      firstName: "Benjamin",
      middleName: "Ocampo",
      clientId: "c2",
      department: "Operations",
      position: "Terminal supervisor",
      basis: "Monthly",
      monthlyRate: 26000,
      dailyRate: 0,
      taxCode: "ME2",
      civilStatus: "Married",
      hired: "2015-03-16",
    }),
    emp({
      id: "e2003",
      empNo: "2003",
      lastName: "Lim",
      firstName: "Carlo",
      middleName: "Uy",
      clientId: "c2",
      department: "Operations",
      position: "Baggage handler",
      dailyRate: 610,
      hired: "2023-02-20",
    }),
    emp({
      id: "e2004",
      empNo: "2004",
      lastName: "Domingo",
      firstName: "Rosa",
      middleName: "Valdez",
      clientId: "c2",
      department: "Operations",
      position: "Ticket teller",
      gender: "Female",
      dailyRate: 645,
      taxCode: "ME4",
      civilStatus: "Married",
      hired: "2020-10-05",
      contractEnd: "2025-12-20",
    }),
    emp({
      id: "e2005",
      empNo: "2005",
      lastName: "Pascual",
      firstName: "Efren",
      middleName: "Gomez",
      clientId: "c2",
      department: "Operations",
      position: "Janitor",
      payrollStatus: "X",
      active: false,
    }),
    emp({
      id: "e3001",
      empNo: "3001",
      lastName: "Morales",
      firstName: "Rodel",
      middleName: "Aguilar",
      clientId: "c3",
      department: "Fleet",
      position: "Bus driver",
      role: "Driver",
      dailyRate: 650,
      shift: "TRIP",
      restDay: "",
      hired: "2016-01-11",
    }),
    emp({
      id: "e3002",
      empNo: "3002",
      lastName: "Salazar",
      firstName: "Nestor",
      middleName: "Pineda",
      clientId: "c3",
      department: "Fleet",
      position: "Bus conductor",
      role: "Conductor",
      dailyRate: 610,
      shift: "TRIP",
      restDay: "",
      hired: "2019-07-22",
    }),
    emp({
      id: "e3003",
      empNo: "3003",
      lastName: "Gonzales",
      firstName: "Arnel",
      middleName: "Reyes",
      clientId: "c3",
      department: "Fleet",
      position: "Bus driver",
      role: "Driver",
      dailyRate: 650,
      shift: "TRIP",
      restDay: "",
      hired: "2018-05-14",
      taxCode: "ME1",
    }),
    emp({
      id: "e3004",
      empNo: "3004",
      lastName: "Rivera",
      firstName: "Joel",
      middleName: "Manalastas",
      clientId: "c3",
      department: "Fleet",
      position: "Bus conductor",
      role: "Conductor",
      dailyRate: 610,
      shift: "TRIP",
      restDay: "",
      hired: "2022-09-01",
    }),
    emp({
      id: "e3005",
      empNo: "3005",
      lastName: "De Leon",
      firstName: "Ricardo",
      middleName: "Sison",
      clientId: "c3",
      department: "Fleet",
      position: "Bus driver",
      role: "Driver",
      dailyRate: 650,
      shift: "TRIP",
      restDay: "",
      hired: "2014-12-01",
    }),
    emp({
      id: "e3006",
      empNo: "3006",
      lastName: "Cruz",
      firstName: "Danilo",
      middleName: "Ventura",
      clientId: "c3",
      department: "Fleet",
      position: "Bus conductor",
      role: "Conductor",
      dailyRate: 610,
      shift: "TRIP",
      restDay: "",
      hired: "2021-04-19",
    }),
    emp({
      id: "e3007",
      empNo: "3007",
      lastName: "Manalo",
      firstName: "Edwin",
      middleName: "Lacson",
      clientId: "c3",
      department: "Fleet",
      position: "Bus driver",
      role: "Driver",
      dailyRate: 650,
      shift: "TRIP",
      restDay: "",
      payrollStatus: "Z",
      active: false,
      contractEnd: "2025-10-31",
    }),
    emp({
      id: "e3008",
      empNo: "3008",
      lastName: "Ocampo",
      firstName: "Victor",
      middleName: "Santos",
      clientId: "c3",
      department: "Fleet",
      position: "Bus driver",
      role: "Driver",
      dailyRate: 650,
      shift: "TRIP",
      restDay: "",
      hired: "2025-08-04",
    }),
  ];

  // Punches for November 2025.
  const punches: Punch[] = [];
  let pid = 0;
  const push = (
    dtrId: string,
    date: string,
    time: string,
    dir: "In" | "Out",
    source = "Terminal TDX-01",
  ) => punches.push({ id: `p${++pid}`, dtrId, date, time, dir, source });
  const nov = dateRange("2025-11-01", "2025-11-30");
  for (const e of employees.filter((x) => x.clientId !== "c3" && !isExcluded(x))) {
    const r = rng(Number(e.empNo) * 7919);
    for (const date of nov) {
      const wd = weekdayOf(date);
      if (wd === e.restDay && !(e.id === "e1007")) continue;
      if (date === "2025-11-01" && e.id !== "e1001") continue; // special holiday, office closed
      if (e.id === "e1004" && (date === "2025-11-20" || date === "2025-11-21")) continue;
      if (e.id === "e1007" && wd === "Mon") continue;
      if (e.id === "e1005" && date === "2025-11-24") {
        push(e.dtrId, date, "08:00", "In");
        push(e.dtrId, date, "08:00", "Out");
        continue;
      }
      if (e.shift === "NIGHT") {
        push(e.dtrId, date, t(1310 + Math.floor(r() * 9)), "In");
        push(e.dtrId, date, t(420 + Math.floor(r() * 12)), "Out");
        continue;
      }
      let inMin = 465 + Math.floor(r() * 13);
      let outMin = 1020 + Math.floor(r() * 18);
      if (e.id === "e1003" && date === "2025-11-18") inMin = 502;
      if (e.id === "e1003" && date === "2025-11-25") inMin = 521;
      if (e.id === "e1001" && date === "2025-11-19") outMin = 1175;
      if (e.id === "e1001" && date === "2025-11-26") outMin = 1205;
      if (e.id === "e1010" && date === "2025-11-28") outMin = 960;
      push(e.dtrId, date, t(inMin), "In");
      push(e.dtrId, date, t(outMin), "Out");
    }
  }
  for (const e of employees.filter((x) => x.clientId === "c3" && !isExcluded(x))) {
    const r = rng(Number(e.empNo) * 104729);
    for (const date of nov) {
      if (r() < 0.14) continue;
      const trips = 2 + Math.floor(r() * 3);
      let cur = 270 + Math.floor(r() * 30);
      for (let i = 0; i < trips; i++) {
        const len = 190 + Math.floor(r() * 50);
        push(e.dtrId, date, t(cur), "In", "FMS dispatch");
        push(e.dtrId, date, t(cur + len), "Out", "FMS dispatch");
        cur += len + 25 + Math.floor(r() * 20);
      }
    }
  }

  const sss: SssRow[] = [{ id: "sss0", from: 0, to: 5249.99, ee: 250, er: 510, ec: 10 }];
  for (let from = 5250, i = 1; from < 34750; from += 500, i++) {
    const msc = from + 250;
    sss.push({
      id: `sss${i}`,
      from,
      to: from + 499.99,
      ee: r2(msc * 0.05),
      er: r2(msc * 0.1),
      ec: msc < 15000 ? 10 : 30,
    });
  }
  sss.push({ id: "sssmax", from: 34750, to: 9999999, ee: 1750, er: 3500, ec: 30 });

  const philhealth: PhRow[] = [{ id: "ph0", from: 0, to: 10000, ee: 250, er: 250 }];
  for (let from = 10000.01, i = 1; from < 100000; from += 1000, i++) {
    philhealth.push({
      id: `ph${i}`,
      from: r2(from),
      to: r2(from + 999.99),
      ee: r2(Math.floor(from) * 0.025),
      er: r2(Math.floor(from) * 0.025),
    });
  }
  philhealth.push({ id: "phmax", from: 100000.01, to: 9999999, ee: 2500, er: 2500 });

  const tax: TaxRow[] = [];
  const brackets2018: [number, number, number][] = [
    [0, 0, 0],
    [10417, 15, 0],
    [16667, 20, 937.5],
    [33333, 25, 4270.7],
    [83333, 30, 16770.7],
    [333333, 35, 91770.7],
  ];
  brackets2018.forEach(([floor, percent, baseTax], i) =>
    tax.push({ id: `t18-${i}`, table: "2018", taxCode: "All codes", floor, percent, baseTax }),
  );
  const oldSteps: [number, number, number][] = [
    [0, 0, 0],
    [0, 5, 0],
    [417, 10, 20.83],
    [1250, 15, 104.17],
    [2917, 20, 354.17],
    [5833, 25, 937.5],
    [10417, 30, 2083.33],
    [20833, 32, 5208.33],
  ];
  const exemptions: Record<string, number> = {
    S: 2083,
    ME: 2083,
    S1: 3125,
    ME1: 3125,
    S2: 4167,
    ME2: 4167,
    ME4: 6250,
  };
  Object.entries(exemptions).forEach(([code, ex]) =>
    oldSteps.forEach(([off, percent, baseTax], i) =>
      tax.push({
        id: `tsm-${code}-${i}`,
        table: "Semi-monthly",
        taxCode: code,
        floor: i === 0 ? 0 : ex + off,
        percent,
        baseTax,
      }),
    ),
  );

  const trips: Trip[] = [];
  const routes = ["Cubao – Batangas City", "Cubao – Lipa", "Pasay – Nasugbu", "Cubao – Lucena"];
  employees
    .filter((e) => e.clientId === "c3" && !isExcluded(e))
    .forEach((e) => {
      const r = rng(Number(e.empNo) * 31337);
      for (const [y, m] of [
        [2025, 10],
        [2025, 11],
      ]) {
        if (e.id === "e3008" && m === 11) continue;
        const dim = new Date(y, m, 0).getDate();
        for (let d = 1; d <= dim; d++) {
          if (r() < 0.3) continue;
          const commission = r2((e.role === "Driver" ? 520 : 380) + r() * 160);
          trips.push({
            id: `tr-${e.id}-${m}-${d}`,
            db: "fms_trips_2025",
            empId: e.id,
            date: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`,
            bundle: `B-${10200 + Math.floor(r() * 800)}`,
            commission,
            fund: r2(commission * 0.02),
            bus: `TAS-${110 + Math.floor(r() * 30)}`,
            route: routes[Math.floor(r() * routes.length)],
          });
        }
      }
    });

  const db: Db = {
    agencies: [
      { code: "TAS01", name: "TAS Trans Corporation", database: "tas_payroll" },
      { code: "TAS02", name: "TAS Express Lines", database: "tas_express" },
      { code: "TAS03", name: "TAS Logistics", database: null },
    ],
    users,
    company: null,
    clients,
    employees,
    shifts: [
      {
        id: "s1",
        code: "DAY8",
        start: "08:00",
        end: "17:00",
        midday: "12:00",
        workHours: 8,
        graceMinutes: 10,
        graceOn: true,
        noLate: false,
        noUndertime: false,
        undertimeMinutes: 0,
        noInOut: false,
        flexible: false,
        noOvertime: false,
        overtimeStart: "17:30",
        nightDiff: false,
      },
      {
        id: "s2",
        code: "NIGHT",
        start: "22:00",
        end: "07:00",
        midday: "02:00",
        workHours: 8,
        graceMinutes: 10,
        graceOn: true,
        noLate: false,
        noUndertime: false,
        undertimeMinutes: 0,
        noInOut: false,
        flexible: false,
        noOvertime: false,
        overtimeStart: "07:30",
        nightDiff: true,
      },
      {
        id: "s3",
        code: "FLEX",
        start: "07:00",
        end: "16:00",
        midday: "12:00",
        workHours: 8,
        graceMinutes: 0,
        graceOn: false,
        noLate: true,
        noUndertime: false,
        undertimeMinutes: 15,
        noInOut: false,
        flexible: true,
        noOvertime: true,
        overtimeStart: "",
        nightDiff: false,
      },
      {
        id: "s4",
        code: "TRIP",
        start: "04:00",
        end: "20:00",
        midday: "12:00",
        workHours: 8,
        graceMinutes: 0,
        graceOn: false,
        noLate: true,
        noUndertime: true,
        undertimeMinutes: 0,
        noInOut: false,
        flexible: true,
        noOvertime: true,
        overtimeStart: "",
        nightDiff: false,
      },
    ],
    holidays: [
      { id: "h1", date: "2025-08-21", name: "Ninoy Aquino Day", letter: "H" },
      { id: "h2", date: "2025-08-25", name: "National Heroes Day", letter: "L" },
      { id: "h3", date: "2025-10-31", name: "All Saints' Day Eve", letter: "H" },
      { id: "h4", date: "2025-11-01", name: "All Saints' Day", letter: "H" },
      { id: "h5", date: "2025-11-30", name: "Bonifacio Day", letter: "L" },
      { id: "h6", date: "2025-12-08", name: "Feast of the Immaculate Conception", letter: "H" },
      { id: "h7", date: "2025-12-24", name: "Christmas Eve", letter: "H" },
      { id: "h8", date: "2025-12-25", name: "Christmas Day", letter: "L" },
      { id: "h9", date: "2025-12-30", name: "Rizal Day", letter: "L" },
      { id: "h10", date: "2025-12-31", name: "Last day of the year", letter: "H" },
    ],
    leaves: [
      { id: "l1", empId: "e1004", date: "2025-11-21", letter: "M", hours: 8 },
      { id: "l2", empId: "e1002", date: "2025-11-27", letter: "V", hours: 8 },
      { id: "l3", empId: "e2001", date: "2025-11-06", letter: "V", hours: 8 },
      { id: "l4", empId: "e2002", date: "2025-11-10", letter: "B", hours: 4 },
      { id: "l5", empId: "e3001", date: "2025-11-30", letter: "D", hours: 8 },
      { id: "l6", empId: "e1003", date: "2025-11-12", letter: "S", hours: 8 },
    ],
    overtimes: [{ id: "o1", empId: "e1001", date: "2025-11-26", hours: 2 }],
    punches,
    staffFinal: {},
    fmsFinal: {},
    payrollAttendance: {},
    cutoffs: [],
    deductions: [
      {
        id: "d1",
        empId: "e1001",
        code: 3,
        balance: 3000,
        schedule: "Every cutoff",
        amount: 500,
        locked: false,
      },
      {
        id: "d2",
        empId: "e1001",
        code: 5,
        balance: 12450,
        schedule: "Second half",
        amount: 1037.5,
        locked: true,
      },
      {
        id: "d3",
        empId: "e1003",
        code: 1,
        balance: 1200,
        schedule: "Every cutoff",
        amount: 300,
        locked: false,
      },
      {
        id: "d4",
        empId: "e1004",
        code: 6,
        balance: 8600,
        schedule: "First half",
        amount: 716.67,
        locked: false,
      },
      {
        id: "d5",
        empId: "e1005",
        code: 7,
        balance: 2000,
        schedule: "Every cutoff",
        amount: 500,
        locked: false,
      },
      {
        id: "d6",
        empId: "e1010",
        code: 4,
        balance: 600,
        schedule: "Second half",
        amount: 150,
        locked: false,
      },
      {
        id: "d7",
        empId: "e3001",
        code: 4,
        balance: 5000,
        schedule: "Every cutoff",
        amount: 1000,
        locked: false,
      },
      {
        id: "d8",
        empId: "e3002",
        code: 8,
        balance: 450,
        schedule: "Every cutoff",
        amount: 450,
        locked: false,
      },
      {
        id: "d9",
        empId: "e3001",
        code: 12,
        balance: 6000,
        schedule: "Every cutoff",
        amount: 100,
        locked: true,
      },
      {
        id: "d10",
        empId: "e2003",
        code: 17,
        balance: 9800,
        schedule: "Second half",
        amount: 816.67,
        locked: false,
      },
    ],
    openPayroll: [],
    history: [],
    trips,
    payments: [],
    thirteenth: [],
    banks: [
      { id: "b1", code: "BDO", name: "BDO Unibank" },
      { id: "b2", code: "BPI", name: "Bank of the Philippine Islands" },
      { id: "b3", code: "LBP", name: "Land Bank of the Philippines" },
      { id: "b4", code: "MBTC", name: "Metrobank" },
    ],
    branches: [
      { id: "br1", code: "CUB", name: "Cubao Araneta", bank: "BDO" },
      { id: "br2", code: "BTG", name: "Batangas Rizal Ave", bank: "BDO" },
      { id: "br3", code: "QA", name: "Quezon Avenue", bank: "BPI" },
    ],
    departments: [
      { id: "dp1", code: "ACC", name: "Accounting" },
      { id: "dp2", code: "HR", name: "Human resources" },
      { id: "dp3", code: "OPS", name: "Operations" },
      { id: "dp4", code: "MNT", name: "Maintenance" },
      { id: "dp5", code: "SEC", name: "Security" },
      { id: "dp6", code: "FLT", name: "Fleet" },
    ],
    sss,
    philhealth,
    hdmf: [
      { id: "hd1", from: 0, to: 1500, eeRate: 1, erRate: 2 },
      { id: "hd2", from: 1500.01, to: 9999999, eeRate: 2, erRate: 2 },
    ],
    tax,
    taxCodes: [
      { id: "tc1", code: "S", description: "Single", exemption: 50000 },
      { id: "tc2", code: "ME", description: "Married", exemption: 50000 },
      { id: "tc3", code: "S1", description: "Single, 1 dependent", exemption: 75000 },
      { id: "tc4", code: "ME1", description: "Married, 1 dependent", exemption: 75000 },
      { id: "tc5", code: "S2", description: "Single, 2 dependents", exemption: 100000 },
      { id: "tc6", code: "ME2", description: "Married, 2 dependents", exemption: 100000 },
      { id: "tc7", code: "ME4", description: "Married, 4 dependents", exemption: 150000 },
    ],
    logs: [],
    benefits: [
      {
        id: "bn1",
        empId: "e1001",
        benefit: "Rice subsidy",
        amount: 1500,
        start: "2025-01-01",
        end: "2025-12-31",
      },
      {
        id: "bn2",
        empId: "e1002",
        benefit: "Medical allowance",
        amount: 2000,
        start: "2025-01-01",
        end: "2025-12-31",
      },
    ],
    assignments: [
      {
        id: "as1",
        empId: "e1001",
        client: "TAS Head Office",
        position: "Accounting clerk",
        start: "2019-06-17",
        end: "",
      },
      {
        id: "as2",
        empId: "e1003",
        client: "Cubao Terminal",
        position: "Dispatcher",
        start: "2018-09-10",
        end: "2021-12-31",
      },
      {
        id: "as3",
        empId: "e1003",
        client: "TAS Head Office",
        position: "Dispatcher",
        start: "2022-01-01",
        end: "",
      },
    ],
    trainings: [
      {
        id: "tn1",
        empId: "e1001",
        title: "Basic bookkeeping",
        provider: "TESDA",
        date: "2024-03-12",
        hours: 16,
      },
    ],
    offenses: [
      {
        id: "of1",
        empId: "e1003",
        date: "2025-07-08",
        offense: "Tardiness, third instance",
        action: "Written warning",
      },
    ],
    leaveBalances: [
      { id: "lb1", empId: "e1001", year: 2025, type: "Sick leave", entitled: 5, used: 1 },
      { id: "lb2", empId: "e1001", year: 2025, type: "Vacation leave", entitled: 5, used: 2 },
      { id: "lb3", empId: "e1004", year: 2025, type: "Sick leave", entitled: 5, used: 3 },
      { id: "lb4", empId: "e1002", year: 2025, type: "Vacation leave", entitled: 10, used: 4 },
    ],
    appraisals: [
      {
        id: "ap1",
        empId: "e1001",
        date: "2025-06-30",
        rating: "Very satisfactory",
        rater: "Carmela Bautista",
        remarks: "Accurate, on time with reconciliations.",
      },
    ],
    session: null,
    flash: null,
  };

  /* Run cutoffs the way the office would have, so every state is visible. */
  const cutoffRec = (c: Cutoff, billed: boolean, by: string | null = null): CutoffRecord => ({
    ...c,
    key: cutoffKey(c.end),
    billed,
    billedAt: billed ? `${c.end}T18:00:00` : null,
    billedBy: by,
  });
  const finalize = (e: Employee, c: Cutoff) => {
    const key = cutoffKey(c.end);
    const fms = e.clientId === "c3";
    if (fms) {
      const rows = fmsBatch(db, e, c);
      db.fmsFinal[`${e.id}|${key}`] = rows;
      db.payrollAttendance[`${e.id}|${key}`] = { kind: "fms", fms: rows, at: `${c.end}T16:00:00` };
    } else {
      const rows = staffBatch(db, e, c);
      db.staffFinal[`${e.id}|${key}`] = rows;
      db.payrollAttendance[`${e.id}|${key}`] = {
        kind: "staff",
        staff: rows,
        at: `${c.end}T16:00:00`,
      };
    }
  };
  const pay = (e: Employee, c: Cutoff, user: string) => {
    const client = clients.find((x) => x.id === e.clientId)!;
    const att = attendanceTotals(db, e, client, cutoffKey(c.end));
    if (!att) return null;
    const { record, problems } = computePayroll(db, e, client, c, att, emptyInputs(), user);
    return problems.length ? null : { ...record, savedAt: `${c.end}T17:00:00` };
  };
  const included = (cid: string) => employees.filter((e) => e.clientId === cid && !isExcluded(e));

  // Head office, Nov 1–15: closed into history.
  const ho1 = { clientId: "c1", start: "2025-11-01", end: "2025-11-15" };
  for (const e of included("c1")) {
    finalize(e, ho1);
    const rec = pay(e, ho1, "staffpay");
    if (rec) db.history.push({ ...rec, movedAt: "2025-11-17T10:12:00" });
    delete db.payrollAttendance[`${e.id}|${cutoffKey(ho1.end)}`];
    delete db.staffFinal[`${e.id}|${cutoffKey(ho1.end)}`];
  }
  // Head office, Nov 16–30: open, partly processed.
  const ho2 = { clientId: "c1", start: "2025-11-16", end: "2025-11-30" };
  for (const e of included("c1")) {
    if (e.id === "e1009") continue;
    finalize(e, ho2);
    if (e.id === "e1007") delete db.payrollAttendance[`${e.id}|${cutoffKey(ho2.end)}`];
  }
  for (const id of ["e1001", "e1002", "e1006"]) {
    const rec = pay(
      employees.find((e) => e.id === id)!,
      ho2,
      "staffpay",
    );
    if (rec) db.openPayroll.push(rec);
  }
  // Fleet, Nov 1–15 closed; Nov 16–30 open.
  const f1 = { clientId: "c3", start: "2025-11-01", end: "2025-11-15" };
  for (const e of included("c3")) {
    finalize(e, f1);
    const rec = pay(e, f1, "fmspay");
    if (rec) db.history.push({ ...rec, movedAt: "2025-11-17T15:40:00" });
    delete db.payrollAttendance[`${e.id}|${cutoffKey(f1.end)}`];
    delete db.fmsFinal[`${e.id}|${cutoffKey(f1.end)}`];
  }
  const f2 = { clientId: "c3", start: "2025-11-16", end: "2025-11-30" };
  for (const id of ["e3001", "e3002", "e3003"])
    finalize(
      employees.find((e) => e.id === id)!,
      f2,
    );
  const fr = pay(
    employees.find((e) => e.id === "e3001")!,
    f2,
    "fmspay",
  );
  if (fr) db.openPayroll.push(fr);

  db.cutoffs = [
    cutoffRec({ clientId: "c1", start: "2025-10-16", end: "2025-10-31" }, true, "staffpay"),
    cutoffRec(ho1, true, "staffpay"),
    cutoffRec(ho2, false),
    cutoffRec({ clientId: "c2", start: "2025-10-16", end: "2025-10-31" }, true, "staffpay"),
    cutoffRec({ clientId: "c2", start: "2025-11-01", end: "2025-11-15" }, false),
    cutoffRec(f1, true, "fmspay"),
    cutoffRec(f2, false),
  ];

  // Stored 13th-month amounts per year.
  for (const e of employees.filter((x) => !isExcluded(x))) {
    const client = clients.find((c) => c.id === e.clientId)!;
    const { daily } = rateBook(e, client);
    db.thirteenth.push({
      id: `13-${e.id}-2024`,
      empId: e.id,
      clientId: e.clientId,
      year: 2024,
      amount: r2((daily * 302) / 12),
    });
    db.thirteenth.push({
      id: `13-${e.id}-2025`,
      empId: e.id,
      clientId: e.clientId,
      year: 2025,
      amount: r2((daily * 288) / 12),
    });
  }

  // Saved commission payments for October.
  for (const id of ["e3001", "e3003"]) {
    const commission = r2(
      trips
        .filter((x) => x.empId === id && x.date.startsWith("2025-10"))
        .reduce((a, x) => a + x.commission, 0),
    );
    const c = commissionContributions(db, commission, 20000);
    db.payments.push({
      id: `pay-${id}-10`,
      empId: id,
      year: 2025,
      month: 10,
      commission,
      suretyFund: r2(
        trips
          .filter((x) => x.empId === id && x.date.startsWith("2025-10"))
          .reduce((a, x) => a + x.fund, 0),
      ),
      ceiling: 20000,
      beginningBalance: 0,
      taxAmount: 200,
      sssEe: c.sssEe,
      sssEr: c.sssEr,
      phEe: c.phEe,
      phEr: c.phEr,
      hdmfEe: c.hdmfEe,
      hdmfEr: c.hdmfEr,
      ecc: c.ecc,
      contributionsUpdated: true,
      savedBy: "fmspay",
      savedAt: "2025-11-05T11:20:00",
    });
  }

  db.logs = [
    {
      id: "lg1",
      at: "2025-11-17T10:12:00",
      user: "staffpay",
      action: "Move to master",
      record: "TAS Head Office, cutoff 20251101",
    },
    {
      id: "lg2",
      at: "2025-11-17T15:40:00",
      user: "fmspay",
      action: "Move to master",
      record: "TAS Fleet Drivers and Conductors, cutoff 20251101",
    },
    {
      id: "lg3",
      at: "2025-11-28T09:15:00",
      user: "jbernardo",
      action: "Edit employee",
      record: "1010 Ramos, Teresa C.: contract end",
    },
    {
      id: "lg4",
      at: "2025-12-01T16:40:00",
      user: "staffpay",
      action: "Save payroll",
      record: "1001 Dela Cruz, Juan P., cutoff 20251102",
    },
    {
      id: "lg5",
      at: "2025-12-01T16:52:00",
      user: "staffpay",
      action: "Save payroll",
      record: "1002 Santos, Maria L., cutoff 20251102",
    },
    { id: "lg6", at: "2025-12-02T07:48:00", user: "jbernardo", action: "Log in", record: "TAS01" },
  ];

  return db;
}
