import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { after, before, test } from "node:test";
import { fileURLToPath } from "node:url";
import { createServer } from "vite";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
let vite;
let engine;
let rules;
let actions;
let desk;

before(async () => {
  vite = await createServer({
    root,
    configFile: false,
    server: { middlewareMode: true },
    appType: "custom",
    logLevel: "error",
  });
  engine = await vite.ssrLoadModule("/src/server/engine.ts");
  rules = await vite.ssrLoadModule("/src/data/rules.ts");
  actions = await vite.ssrLoadModule("/src/server/actions.ts");
  desk = await vite.ssrLoadModule("/src/server/desk.ts");
});

after(async () => {
  await vite?.close();
});

const rates = {
  overtime: 1.25,
  nightDiff: 0.1,
  restDay: 0.3,
  specialHoliday: 0.3,
  legalHoliday: 2,
  holidayOvertime: 0,
  excessOvertime: 0,
  ecola: 0,
  hazard: 0,
  agencyRate: 0,
  dailyCharge: 0,
};

function client(endDayFlags) {
  return {
    id: "c1",
    code: "C1",
    code2: "",
    name: "Client",
    active: true,
    workingDays: 303,
    rates,
    flags: {
      sss: true,
      philhealth: true,
      pagibig: true,
      thirteenthOnCutoff: true,
      displayThirteenth: true,
      tax2018: false,
      ...endDayFlags,
    },
    taxPercent: false,
    taxCeiling: 0,
  };
}

function employee() {
  return {
    id: "e1",
    empNo: "1001",
    dtrId: "D1",
    lastName: "Dela Cruz",
    firstName: "Juan",
    middleName: "",
    department: "Ops",
    position: "Clerk",
    role: "",
    taxCode: "S",
    civilStatus: "Single",
    gender: "Male",
    birthDate: "",
    address: "",
    phone: "",
    tin: "",
    sssNo: "",
    philhealthNo: "",
    pagibigNo: "",
    bank: "",
    accountNo: "",
    basis: "Hourly",
    monthlyRate: 0,
    dailyRate: 600,
    hourlyRate: 0,
    sss: true,
    philhealth: true,
    pagibig: true,
    philhealthPerDay: 0,
    pagibigPerDay: 0,
    shift: "A",
    restDay: "Sun",
    clientId: "c1",
    hired: "",
    contractStart: "",
    contractEnd: "",
    payrollStatus: "A",
    active: true,
  };
}

function db(extra = {}) {
  return {
    agencies: [],
    users: [],
    company: null,
    clients: [],
    employees: [],
    shifts: [
      {
        id: "s1",
        code: "A",
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
        overtimeStart: "17:00",
        nightDiff: false,
      },
    ],
    holidays: [],
    leaves: [],
    overtimes: [],
    punches: [],
    staffFinal: {},
    fmsFinal: {},
    payrollAttendance: {},
    cutoffs: [],
    deductions: [],
    openPayroll: [],
    history: [],
    trips: [],
    payments: [],
    thirteenth: [],
    banks: [],
    branches: [],
    departments: [],
    sss: [],
    philhealth: [],
    hdmf: [],
    tax: [],
    taxCodes: [],
    logs: [],
    benefits: [],
    assignments: [],
    trainings: [],
    offenses: [],
    leaveBalances: [],
    appraisals: [],
    session: null,
    flash: null,
    ...extra,
  };
}

function attendance(days) {
  return {
    kind: "staff",
    hoursWorked: 80,
    days,
    overtime: 0,
    nd: 0,
    restDay: 0,
    special: 0,
    legal: 0,
    sickDays: 0,
    vacationDays: 0,
    absentDays: 0,
    lateHours: 0,
    absenceAmount: 0,
    lateAmount: 0,
  };
}

function pay(end, days, attExtra = {}) {
  const emp = employee();
  const c = client();
  const att = { ...attendance(days), ...attExtra };
  return engine.computePayroll(db(), emp, c, { clientId: "c1", start: "2025-11-01", end }, att, engine.emptyInputs(), "admin");
}

test("staff required hours are forced to 8", () => {
  const emp = employee();
  const book = db({
    punches: [
      { id: "p1", dtrId: "D1", date: "2025-11-03", time: "08:00", dir: "In", source: "file" },
      { id: "p2", dtrId: "D1", date: "2025-11-03", time: "12:00", dir: "Out", source: "file" },
    ],
  });
  const [day] = engine.staffBatch(book, emp, { clientId: "c1", start: "2025-11-03", end: "2025-11-03" });
  assert.equal(day.required, 8);
  assert.ok(day.work < 8);
});

test("H pays the special holiday rate and L pays the legal holiday rate", () => {
  const emp = employee();
  const book = db({
    payrollAttendance: {
      "e1|20251101": {
        kind: "staff",
        at: "2025-11-15T12:00:00",
        staff: [
          { date: "2025-11-01", weekday: "Sat", timeIn: "", timeOut: "", required: 8, work: 0, late: 0, undertime: 0, overtime: 0, nd: 0, leaveLetter: "H", leaveHours: 8, remark: "Present", manual: false },
          { date: "2025-11-02", weekday: "Sun", timeIn: "", timeOut: "", required: 8, work: 0, late: 0, undertime: 0, overtime: 0, nd: 0, leaveLetter: "L", leaveHours: 8, remark: "Present", manual: false },
        ],
      },
    },
  });
  const totals = engine.attendanceTotals(book, emp, client(), "20251101");
  assert.equal(totals.special, 8);
  assert.equal(totals.legal, 8);
  const { record } = engine.computePayroll(
    book,
    emp,
    client(),
    { clientId: "c1", start: "2025-11-01", end: "2025-11-15" },
    totals,
    engine.emptyInputs(),
    "admin",
  );
  const special = record.earnings.find((line) => line.key === "special");
  const legal = record.earnings.find((line) => line.key === "legal");
  assert.equal(special.amount, 180);
  assert.equal(legal.amount, 1200);
});

test("cutoff end day picks SSS, PhilHealth, Pag-IBIG, tax, and union dues", () => {
  const first = pay("2025-11-15", 10).record.deductions.map((line) => line.key);
  const second = pay("2025-11-30", 10).record.deductions.map((line) => line.key);
  assert.ok(first.includes("sss"));
  assert.ok(first.includes("philhealth"));
  assert.ok(first.includes("tax"));
  assert.ok(first.includes("union"));
  assert.equal(first.includes("pagibig"), false);
  assert.ok(second.includes("pagibig"));
  assert.ok(second.includes("tax"));
  assert.equal(second.includes("sss"), false);
  assert.equal(second.includes("philhealth"), false);
  assert.equal(second.includes("union"), false);
});

test("cutoff 13th month is daily rate divided by 12", () => {
  const { record } = pay("2025-11-15", 10);
  const line = record.earnings.find((item) => item.key === "thirteenth");
  assert.equal(line.amount, 50);
});

test("driver overtime stays 0 when the pairs sum past 8", () => {
  const book = db({
    punches: [
      { id: "a", dtrId: "D1", date: "2025-06-02", time: "06:00", dir: "In", source: "file" },
      { id: "b", dtrId: "D1", date: "2025-06-02", time: "16:00", dir: "Out", source: "file" },
    ],
  });
  const rows = engine.fmsBatch(book, employee(), {
    clientId: "c1",
    start: "2025-06-02",
    end: "2025-06-02",
  });
  const day = rows.find((row) => row.date === "2025-06-02");
  assert.ok(day.required > 8);
  assert.equal(day.overtime, 0);
});

test("February does not create a real day 29", () => {
  const rows = engine.fmsBatch(db(), employee(), {
    clientId: "c1",
    start: "2024-02-01",
    end: "2024-02-29",
  });
  assert.equal(
    rows.some((row) => row.date === "2024-02-29"),
    false,
  );
  const day29 = rows.find((row) => row.day === 29);
  if (day29) assert.equal(day29.remark, "Invalid Day");
});

test("a ceiling of 20000 cuts commission before the statutory brackets", () => {
  const book = db({
    sss: [
      { id: "wide", from: 0, to: 999999, ee: 11, er: 1, ec: 1 },
      { id: "cap", from: 20000, to: 30000, ee: 44, er: 4, ec: 4 },
    ],
    philhealth: [
      { id: "wide", from: 0, to: 999999, ee: 5, er: 5 },
      { id: "cap", from: 20000, to: 30000, ee: 21, er: 21 },
    ],
    hdmf: [{ id: "hd", from: 0, to: 999999, eeRate: 2, erRate: 2 }],
  });
  const shares = engine.commissionContributions(book, 25000, 20000);
  assert.equal(shares.capped, 20000);
  assert.equal(shares.sssEe, 44);
  assert.equal(shares.phEe, 21);
  assert.equal(shares.hdmfEe, 400);
});

test("commission of 0 skips the contribution update and stores no tax", async () => {
  const emp = employee();
  desk.bind(
    db({
      users: [{ id: "u1", username: "admin" }],
      employees: [emp],
      session: { userId: "u1", agencyCode: "TAS01", loginAt: "2025-11-01T00:00:00", cutoff: null },
    }),
  );
  const payment = await actions.saveCommissionPayment({
    empId: emp.id,
    year: 2025,
    month: 11,
    commission: 0,
    suretyFund: 10,
    ceiling: 20000,
    beginningBalance: 0,
    taxAmount: 200,
    taxCode: "S",
  });
  assert.equal(payment.contributionsUpdated, false);
  assert.equal(payment.sssEe, 0);
  assert.equal(payment.taxAmount, 0);
});

test("driver tax stays the entered 200 and ignores the tax table", async () => {
  const emp = employee();
  desk.bind(
    db({
      users: [{ id: "u1", username: "admin" }],
      employees: [emp],
      session: { userId: "u1", agencyCode: "TAS01", loginAt: "2025-11-01T00:00:00", cutoff: null },
      tax: [{ id: "t1", table: "2018", taxCode: "", floor: 0, percent: 50, baseTax: 9999 }],
      sss: [{ id: "s1", from: 0, to: 999999, ee: 1, er: 1, ec: 1 }],
      philhealth: [{ id: "p1", from: 0, to: 999999, ee: 1, er: 1 }],
      hdmf: [{ id: "h1", from: 0, to: 999999, eeRate: 1, erRate: 1 }],
    }),
  );
  const payment = await actions.saveCommissionPayment({
    empId: emp.id,
    year: 2025,
    month: 11,
    commission: 8000,
    suretyFund: 0,
    ceiling: 20000,
    beginningBalance: 0,
    taxAmount: 200,
    taxCode: "S",
  });
  assert.equal(payment.taxAmount, 200);
  assert.equal(payment.contributionsUpdated, true);
});

test("payroll statuses T W X Y Z are excluded and A and P are not", () => {
  for (const status of ["T", "W", "X", "Y", "Z"]) {
    assert.equal(rules.isExcluded({ payrollStatus: status }), true);
  }
  assert.equal(rules.isExcluded({ payrollStatus: "A" }), false);
  assert.equal(rules.isExcluded({ payrollStatus: "P" }), false);
});

test("monthly 13th month uses the employee daily rate divided by 12", () => {
  const emp = employee();
  emp.basis = "Monthly";
  emp.monthlyRate = 30000;
  emp.dailyRate = 610;
  const c = client();
  c.workingDays = 300;
  const { record } = engine.computePayroll(
    db(),
    emp,
    c,
    { clientId: "c1", start: "2025-11-01", end: "2025-11-15" },
    attendance(10),
    engine.emptyInputs(),
    "admin",
  );
  const line = record.earnings.find((item) => item.key === "thirteenth");
  assert.equal(line.amount, 50.83);
});

test("cutoff end day 16 takes SSS and PhilHealth, tax, and no union or Pag-IBIG", () => {
  const keys = pay("2025-11-16", 10).record.deductions.map((line) => line.key);
  assert.ok(keys.includes("sss"));
  assert.ok(keys.includes("philhealth"));
  assert.ok(keys.includes("tax"));
  assert.equal(keys.includes("pagibig"), false);
  assert.equal(keys.includes("union"), false);
});

test("browser screens do not import the pay engine", () => {
  const targets = [];
  for (const folder of ["screens", "ui", "app"]) {
    const dir = path.join(root, "src", folder);
    for (const name of fs.readdirSync(dir)) {
      if (name.endsWith(".ts") || name.endsWith(".tsx")) targets.push(path.join(dir, name));
    }
  }
  targets.push(path.join(root, "src", "data", "api.ts"));
  const hits = targets.filter((file) => fs.readFileSync(file, "utf8").includes("server/engine"));
  assert.deepEqual(
    hits.map((file) => path.relative(root, file)),
    [],
  );
});
