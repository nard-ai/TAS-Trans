import { execFileSync } from "node:child_process";
import { readFileSync, unlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const base = "http://localhost:8443";

function sqlJson(query) {
  const file = join(tmpdir(), `tas-import-${Date.now()}-${Math.random().toString(16).slice(2)}.json`);
  execFileSync(
    "sqlcmd",
    [
      "-S",
      "localhost",
      "-d",
      "HRAndPayroll_ReverseEngineering",
      "-E",
      "-C",
      "-y",
      "0",
      "-o",
      file,
      "-Q",
      `SET NOCOUNT ON; ${query}`,
    ],
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

const r2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
const num = (n) => (n == null || n === "" ? 0 : Number(n));
const day = (v) => (v ? String(v).slice(0, 10) : "");
const trim = (v) => String(v ?? "").trim();
function hm(v) {
  const s = trim(v);
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return "";
  return `${m[1].padStart(2, "0")}:${m[2]}`;
}
function punchTime(v) {
  const clock = hm(v);
  if (clock) return clock;
  const digits = trim(v).match(/^(\d{2})(\d{2})/);
  return digits ? `${digits[1]}:${digits[2]}` : "";
}
function percent(n) {
  const x = num(n);
  return x > 1 ? r2(x) : r2(x * 100);
}
function cutoffKey(end) {
  const [y, m, d] = end.split("-").map(Number);
  return `${y}${String(m).padStart(2, "0")}0${d <= 15 ? 1 : 2}`;
}

const clientsRaw = sqlJson(`
  SELECT ClientCode, ClientName, ClientCode1, ClientCode2, ClientStatus,
         WorkingDaysperYear, OvertimeRate, NDiffRate, RestDayRate,
         SpecialHolidayRate, LegalHolidayRate, SplHolidayOTRate, ExcessOTRegDayRate,
         EcolaRate, HazardPayRate, ClientAgencyRate, ClientChargeperDay,
         ClientSSSFlag, ClientPagibigFlag, ClientPhilHealthFlag,
         With13MonthPay, Display13MonthPay, EcolaPayonRestDay,
         ClientTaxPercentFlag, ClientTaxPercentage
  FROM ClientTable
  FOR JSON PATH`);

const employeesRaw = sqlJson(`
  SELECT e.EmployeeID, e.EmpNo, e.DTRID, e.Surname, e.FirstName, e.Middlename,
         e.EmpPosition, e.TaxCode, e.CivilStatus, e.Gender, e.BirthDate,
         e.Address1, e.City, e.EmpCellphone, e.EmpTelNo, e.TIN, e.SSS, e.Philhealth, e.Pagibig,
         e.AcctNo, e.PaymentBasis, e.MonthlyRate, e.DailyRate, e.HourlyRate,
         e.ShiftCode, e.DayOff, e.ClientCode, e.ClientName, e.DateHired, e.StartContract, e.EndContract,
         e.EmploymentStatus, e.EmpStatus,
         v.WithSSSDeduction, v.WithPHealthDeduction, v.WithPbigDeduction,
         v.PHealthRateperday, v.PagibigRateperDay
  FROM HREmployees e
  OUTER APPLY (
    SELECT TOP 1 WithSSSDeduction, WithPHealthDeduction, WithPbigDeduction,
                 PHealthRateperday, PagibigRateperDay
    FROM vwEmployeeBasicSalary s
    WHERE s.EmployeeID = e.EmployeeID
  ) v
  FOR JSON PATH`);

const shiftsRaw = sqlJson(`
  SELECT ShiftCode, ShiftShort, WorkHours, StartTime, EndTime, MidTime, GracePeriod,
         NoGracePeriod, NoLate, NoUnderTime, UnderTime, NoInOut, FlexiTime, NoOT,
         StartRegOT, NoNightDiff
  FROM HRShiftSchedule
  FOR JSON PATH`);

const holidaysRaw = sqlJson(`
  SELECT HolidayID, LTRIM(RTRIM(HolidayType)) AS letter, HolidayName,
         CONVERT(varchar(10), HolidayDate, 23) AS holidayDate
  FROM tblHolidays
  FOR JSON PATH`);

const sssRaw = sqlJson(`
  SELECT ID, MinWage, MaxWage, EE_Share, ER_Share, EC_Share
  FROM SSSTable ORDER BY MinWage, ID
  FOR JSON PATH`);
const phRaw = sqlJson(`
  SELECT ID, MINWAGE, MAXWAGE, PH_EE, PH_ER FROM PhilhealthTable ORDER BY MINWAGE
  FOR JSON PATH`);
const hdRaw = sqlJson(`
  SELECT ID, MinWage, MaxWage, PagibigEE, PagibigER FROM PagibigTable ORDER BY MinWage
  FOR JSON PATH`);
const taxRaw = sqlJson(`
  SELECT ID, TaxCode, SalRate2, PerOver, BaseTax, Bracket FROM SemiMonthlyTaxTable
  FOR JSON PATH`);
const tax2018Raw = sqlJson(`
  SELECT ID, TaxCode, SalRate2, PerOver, BaseTax, Bracket FROM semimonthlytax2018
  FOR JSON PATH`);
const codesRaw = sqlJson(`
  SELECT ID, TaxExemptionCode AS code, Description, ExemptionAmount FROM TaxExemptionCode
  FOR JSON PATH`);
const cutoffsRaw = sqlJson(`
  SELECT ClientCode, ClientName,
         CONVERT(varchar(10), StartCutoffDate, 23) AS startDate,
         CONVERT(varchar(10), EndCutoffDate, 23) AS endDate,
         BilledFlag, Processedby
  FROM CutoffTable
  WHERE EndCutoffDate >= '2026-01-01' AND EndCutoffDate < '2027-01-01'
  FOR JSON PATH`);
const punchesRaw = sqlJson(`
  SELECT DTRID, CONVERT(varchar(10), DTR_Date, 23) AS punchDate, DTR_Time, DTR_IOMode, Machine_ID
  FROM DTRRecords
  WHERE DTR_Date >= '2026-06-26' AND DTR_Date < '2026-07-11'
  FOR JSON PATH`);
const thirteenthRaw = sqlJson(`
  SELECT EmployeeID, ClientCode, YearOf, Month13Payment,
         CONVERT(varchar(10), EndPayDate, 23) AS endPay
  FROM HR13thMonthPay
  WHERE YearOf >= 2024
  FOR JSON PATH`);

const emptyRates = {
  overtime: 0,
  nightDiff: 0,
  restDay: 0,
  specialHoliday: 0,
  legalHoliday: 0,
  holidayOvertime: 0,
  excessOvertime: 0,
  ecola: 0,
  hazard: 0,
  agencyRate: 0,
  dailyCharge: 0,
};
const clientByCode = new Map();
const clients = [];
function ensureClient(code, name) {
  if (code == null || code === "") return null;
  if (clientByCode.has(code)) return clientByCode.get(code);
  const client = {
    id: `c${code}`,
    code: String(code),
    code2: "",
    name: trim(name) || `Client ${code}`,
    active: false,
    workingDays: null,
    rates: { ...emptyRates },
    flags: {
      sss: false,
      philhealth: false,
      pagibig: false,
      thirteenthOnCutoff: false,
      displayThirteenth: false,
      tax2018: false,
    },
    taxPercent: false,
    taxCeiling: 0,
  };
  clientByCode.set(code, client);
  clients.push(client);
  return client;
}
for (const row of clientsRaw) {
  const client = {
    id: `c${row.ClientCode}`,
    code: trim(row.ClientCode1) || String(row.ClientCode),
    code2: trim(row.ClientCode2),
    name: trim(row.ClientName),
    active: !!row.ClientStatus,
    workingDays: row.WorkingDaysperYear == null ? null : num(row.WorkingDaysperYear),
    rates: {
      overtime: num(row.OvertimeRate),
      nightDiff: num(row.NDiffRate),
      restDay: num(row.RestDayRate),
      specialHoliday: num(row.SpecialHolidayRate),
      legalHoliday: num(row.LegalHolidayRate),
      holidayOvertime: num(row.SplHolidayOTRate),
      excessOvertime: num(row.ExcessOTRegDayRate),
      ecola: num(row.EcolaRate),
      hazard: num(row.HazardPayRate),
      agencyRate: num(row.ClientAgencyRate),
      dailyCharge: num(row.ClientChargeperDay),
    },
    flags: {
      sss: !!row.ClientSSSFlag,
      philhealth: !!row.ClientPhilHealthFlag,
      pagibig: !!row.ClientPagibigFlag,
      thirteenthOnCutoff: !!row.With13MonthPay,
      displayThirteenth: !!row.Display13MonthPay,
      tax2018: !!row.EcolaPayonRestDay,
    },
    taxPercent: !!row.ClientTaxPercentFlag,
    taxCeiling: num(row.ClientTaxPercentage),
  };
  clientByCode.set(row.ClientCode, client);
  clients.push(client);
}

const seenEmpNo = new Set();
const employees = employeesRaw.map((row) => {
  const client = ensureClient(row.ClientCode ?? 0, row.ClientName || "No client");
  let empNo = trim(row.EmpNo) || String(row.EmployeeID);
  if (seenEmpNo.has(empNo)) empNo = `${empNo}-${row.EmployeeID}`;
  seenEmpNo.add(empNo);
  const gender = trim(row.Gender).toUpperCase();
  const position = trim(row.EmpPosition);
  const code2 = client?.code2 ?? "";
  let role = "";
  if (code2 === "FMS") {
    const p = position.toLowerCase();
    if (p.includes("conductor")) role = "Conductor";
    else if (p.includes("driver")) role = "Driver";
  }
  return {
    id: `e${row.EmployeeID}`,
    empNo,
    dtrId: trim(row.DTRID),
    lastName: trim(row.Surname),
    firstName: trim(row.FirstName),
    middleName: trim(row.Middlename),
    department: "",
    position,
    role,
    taxCode: trim(row.TaxCode),
    civilStatus: trim(row.CivilStatus),
    gender: gender === "M" ? "Male" : gender === "F" ? "Female" : gender,
    birthDate: day(row.BirthDate),
    address: [trim(row.Address1), trim(row.City)].filter(Boolean).join(", "),
    phone: trim(row.EmpCellphone) || trim(row.EmpTelNo),
    tin: trim(row.TIN),
    sssNo: trim(row.SSS),
    philhealthNo: trim(row.Philhealth),
    pagibigNo: trim(row.Pagibig),
    bank: "",
    accountNo: trim(row.AcctNo),
    basis: trim(row.PaymentBasis).toLowerCase() === "monthly" ? "Monthly" : "Hourly",
    monthlyRate: num(row.MonthlyRate),
    dailyRate: num(row.DailyRate),
    hourlyRate: num(row.HourlyRate),
    sss: !!row.WithSSSDeduction,
    philhealth: !!row.WithPHealthDeduction,
    pagibig: !!row.WithPbigDeduction,
    philhealthPerDay: num(row.PHealthRateperday),
    pagibigPerDay: num(row.PagibigRateperDay),
    shift: row.ShiftCode == null ? "" : String(row.ShiftCode),
    restDay: trim(row.DayOff),
    clientId: client.id,
    hired: day(row.DateHired),
    contractStart: day(row.StartContract),
    contractEnd: day(row.EndContract),
    payrollStatus: trim(row.EmploymentStatus).toUpperCase() || "A",
    active: !!row.EmpStatus,
  };
});

const shifts = shiftsRaw.map((row) => ({
  id: `s${row.ShiftCode}`,
  code: String(row.ShiftCode),
  start: hm(row.StartTime),
  end: hm(row.EndTime),
  midday: hm(row.MidTime),
  workHours: num(row.WorkHours) || 8,
  graceMinutes: num(row.GracePeriod),
  graceOn: !row.NoGracePeriod,
  noLate: !!row.NoLate,
  noUndertime: !!row.NoUnderTime,
  undertimeMinutes: num(row.UnderTime),
  noInOut: !!row.NoInOut,
  flexible: !!row.FlexiTime,
  noOvertime: !!row.NoOT,
  overtimeStart: hm(row.StartRegOT),
  nightDiff: !!row.NoNightDiff,
}));

const holidays = holidaysRaw
  .filter((row) => (row.letter === "H" || row.letter === "L") && row.holidayDate)
  .map((row) => ({
    id: `h${row.HolidayID}`,
    date: row.holidayDate,
    name: trim(row.HolidayName),
    letter: row.letter,
  }));

const sss = sssRaw.map((row) => ({
  id: `sss${row.ID}`,
  from: num(row.MinWage),
  to: num(row.MaxWage),
  ee: num(row.EE_Share),
  er: num(row.ER_Share),
  ec: num(row.EC_Share),
}));
const philhealth = phRaw.map((row) => ({
  id: `ph${row.ID}`,
  from: num(row.MINWAGE),
  to: num(row.MAXWAGE),
  ee: num(row.PH_EE),
  er: num(row.PH_ER),
}));
const hdmf = hdRaw.map((row) => ({
  id: `hd${row.ID}`,
  from: num(row.MinWage),
  to: num(row.MaxWage),
  eeRate: percent(row.PagibigEE),
  erRate: percent(row.PagibigER),
}));
const tax = [
  ...taxRaw.map((row) => ({
    id: `tsm-${row.ID}`,
    table: "Semi-monthly",
    taxCode: trim(row.TaxCode),
    floor: num(row.SalRate2),
    percent: percent(row.PerOver),
    baseTax: num(row.BaseTax),
  })),
  ...tax2018Raw.map((row) => ({
    id: `t18-${row.ID}`,
    table: "2018",
    taxCode: trim(row.TaxCode) || "All codes",
    floor: num(row.SalRate2),
    percent: percent(row.PerOver),
    baseTax: num(row.BaseTax),
  })),
];
const taxCodes = codesRaw.map((row) => ({
  id: `tc${row.ID}`,
  code: trim(row.code),
  description: trim(row.Description),
  exemption: num(row.ExemptionAmount),
}));

const cutoffMap = new Map();
for (const row of cutoffsRaw) {
  if (!row.endDate || !row.startDate || row.ClientCode == null) continue;
  ensureClient(row.ClientCode, row.ClientName);
  const key = cutoffKey(row.endDate);
  const id = `c${row.ClientCode}|${key}`;
  const prev = cutoffMap.get(id);
  if (prev && prev.end >= row.endDate) continue;
  cutoffMap.set(id, {
    clientId: `c${row.ClientCode}`,
    start: row.startDate,
    end: row.endDate,
    key,
    billed: !!row.BilledFlag,
    billedAt: null,
    billedBy: trim(row.Processedby) || null,
  });
}
const cutoffs = [...cutoffMap.values()];

const punches = [];
for (const row of punchesRaw) {
  const time = punchTime(row.DTR_Time);
  const dtrId = trim(row.DTRID);
  if (!time || !dtrId || !row.punchDate) continue;
  punches.push({
    id: `p${punches.length + 1}`,
    dtrId,
    date: row.punchDate,
    time,
    dir: Number(row.DTR_IOMode) === 1 ? "Out" : "In",
    source: row.Machine_ID == null ? "" : String(row.Machine_ID),
  });
}

const thirteenthMap = new Map();
for (const row of thirteenthRaw) {
  const id = `13-e${row.EmployeeID}-${row.YearOf}`;
  const prev = thirteenthMap.get(id);
  if (prev && prev.endPay >= (row.endPay || "")) continue;
  thirteenthMap.set(id, {
    id,
    empId: `e${row.EmployeeID}`,
    clientId: ensureClient(row.ClientCode ?? 0, "No client").id,
    year: Number(row.YearOf),
    amount: r2(num(row.Month13Payment)),
    endPay: row.endPay || "",
  });
}
const thirteenth = [...thirteenthMap.values()].map(({ endPay, ...row }) => row);

const got = await fetch(`${base}/api/db`).then((r) => r.json());
const db = got.db;
const open = cutoffs.find((c) => !c.billed) ?? cutoffs[cutoffs.length - 1] ?? null;
db.clients = clients;
db.employees = employees;
db.shifts = shifts;
db.holidays = holidays;
db.sss = sss;
db.philhealth = philhealth;
db.hdmf = hdmf;
db.tax = tax;
db.taxCodes = taxCodes;
db.cutoffs = cutoffs;
db.punches = punches;
db.thirteenth = thirteenth;
db.leaves = [];
db.overtimes = [];
db.staffFinal = {};
db.fmsFinal = {};
db.payrollAttendance = {};
db.deductions = [];
db.openPayroll = [];
db.history = [];
db.trips = [];
db.payments = [];
if (db.session) db.session.cutoff = open ? { clientId: open.clientId, start: open.start, end: open.end } : null;
db.logs = [
  ...(db.logs ?? []),
    {
    id: `log-import-${Date.now()}`,
    at: new Date().toISOString(),
    user: "admin",
    action: "Import real payroll",
    record: `${employees.length} employees, ${punches.length} punches from 2026-06-26 to 2026-07-10`,
  },
];

const body = JSON.stringify(db);
console.log(
  `ready clients=${clients.length} employees=${employees.length} punches=${punches.length} cutoffs=${cutoffs.length} bytes=${body.length}`,
);
if (body.length > 7_500_000) throw new Error(`Import is ${body.length} bytes and will not fit one save.`);
const put = await fetch(`${base}/api/db`, {
  method: "PUT",
  headers: { "content-type": "application/json" },
  body,
});
if (!put.ok) throw new Error(await put.text());
console.log("saved");
