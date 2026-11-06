/*
 * Stand-in for the payroll server.
 *
 * In production these rules run on the server against Neon. Screens never import this
 * module directly: they call `src/data/api.ts`, which plays the role of the HTTP layer.
 */
import type {
  AttendanceTotals,
  Client,
  Cutoff,
  Db,
  Employee,
  FmsDay,
  PayInputs,
  PayLine,
  PayrollRecord,
  StaffDay,
} from "../data/types";
import {
  cutoffHalf,
  cutoffKey,
  dateRange,
  daysInMonth,
  deductionLabel,
  isFmsClient,
  isoOf,
  parseISO,
  showsSssPhilhealth,
  showsUnionDues,
  weekdayOf,
} from "../data/rules";
import { emptyInputs, r2, uid } from "../data/money";

export { emptyInputs, r2, uid };
export const toMin = (t: string) => {
  const [h, m] = t.split(":").map(Number);
  return h * 60 + m;
};
function overlap(a: number, b: number, c: number, d: number) {
  return Math.max(0, Math.min(b, d) - Math.max(a, c));
}

/* ---------- Staff batch ---------- */

export function staffBatch(db: Db, emp: Employee, cutoff: Cutoff): StaffDay[] {
  const ctx = shiftContext(db, emp);
  const { shift } = ctx;
  const days = dateRange(cutoff.start, cutoff.end).map<StaffDay>((date) => {
    const weekday = weekdayOf(date);
    const dayOff = emp.restDay === weekday;
    const hol = db.holidays.find((h) => h.date === date);
    const ps = db.punches
      .filter((p) => p.dtrId === emp.dtrId && p.date === date)
      .sort((a, b) => a.time.localeCompare(b.time));
    const ins = ps.filter((p) => p.dir === "In");
    const outs = ps.filter((p) => p.dir === "Out");
    const day: StaffDay = {
      date,
      weekday,
      timeIn: ins[0]?.time ?? "",
      timeOut: outs[outs.length - 1]?.time ?? "",
      required: 0,
      work: 0,
      late: 0,
      undertime: 0,
      overtime: 0,
      nd: 0,
      leaveLetter: "",
      leaveHours: 0,
      remark: "Absent",
      manual: false,
    };
    // Night shift: the out punch after midnight is stored on the shift's start date.
    if (shift.nightDiff && ins.length && outs.length)
      day.timeOut = outs.find((o) => o.time < ins[0].time)?.time ?? day.timeOut;
    evaluateDay(day, ctx);
    if (day.remark === "Absent" && (dayOff || hol)) day.remark = "Day off";
    if (hol) {
      day.leaveLetter = hol.letter;
      day.leaveHours = day.remark === "Present" ? Math.min(8, day.work) : 0;
    } else if (dayOff && day.remark === "Present") {
      day.leaveLetter = "R";
      day.leaveHours = Math.min(8, day.work);
    }
    if (shift.noInOut && !dayOff && day.remark !== "Present") {
      day.remark = "Present";
      day.required = 8;
    }
    return day;
  });
  return applyLeaveAndOvertime(db, emp, days);
}

type ShiftCtx = {
  shift: Db["shifts"][number];
  sStart: number;
  sEnd: number;
  midday: number;
  otStart: number;
  spanHours: number;
};

export function evaluateDay(day: StaffDay, ctx: ShiftCtx) {
  const { shift, sStart, sEnd, midday, otStart, spanHours } = ctx;
  day.required = 0;
  day.work = 0;
  day.late = 0;
  day.undertime = 0;
  day.overtime = 0;
  day.nd = 0;
  if (!day.timeIn && !day.timeOut) {
    day.remark = "Absent";
    return;
  }
  if (!day.timeIn || !day.timeOut || day.timeIn === day.timeOut) {
    day.remark = "No In/Out";
    return;
  }
  let a = toMin(day.timeIn);
  let b = toMin(day.timeOut);
  if (shift.nightDiff && a < 720 && sStart > 720) a += 1440;
  if (b <= a) b += 1440;
  let work = b - a;
  if (spanHours > shift.workHours && a < midday + 60 && b > midday) work -= 60;
  day.work = r2(Math.max(0, work) / 60);
  if (!shift.noLate && !shift.flexible) {
    const lateMin = a - sStart;
    const graced = shift.graceOn && lateMin <= shift.graceMinutes;
    day.late = lateMin > 0 && !graced ? r2(lateMin / 60) : 0;
  }
  if (!shift.noUndertime && !shift.flexible) {
    const ut = sEnd - b;
    day.undertime = ut > shift.undertimeMinutes ? r2(ut / 60) : 0;
  }
  if (!shift.noOvertime && b > otStart) day.overtime = Math.floor((b - sEnd) / 30) * 0.5;
  if (shift.nightDiff)
    day.nd = r2(
      (overlap(a, b, 1320, 1800) + overlap(a, b, 0, 360) + overlap(a, b, 2760, 3240)) / 60,
    );
  // Pay rule: every processed day is forced to 8 required hours.
  day.required = 8;
  day.remark = "Present";
}

export function shiftContext(db: Db, emp: Employee): ShiftCtx {
  const shift = db.shifts.find((s) => s.code === emp.shift) ?? db.shifts[0];
  const sStart = toMin(shift.start);
  let sEnd = toMin(shift.end);
  if (sEnd <= sStart) sEnd += 1440;
  let midday = toMin(shift.midday);
  if (midday < sStart) midday += 1440;
  let otStart = shift.overtimeStart ? toMin(shift.overtimeStart) : sEnd;
  if (otStart < sStart) otStart += 1440;
  return { shift, sStart, sEnd, midday, otStart, spanHours: (sEnd - sStart) / 60 };
}

export function applyLeaveAndOvertime(db: Db, emp: Employee, days: StaffDay[]): StaffDay[] {
  return days.map((d) => {
    const day = { ...d };
    const leave = db.leaves.find((l) => l.empId === emp.id && l.date === d.date);
    if (leave) {
      day.leaveLetter = leave.letter;
      day.leaveHours = leave.hours;
      if (leave.letter === "B") {
        day.required = leave.hours;
        day.remark = "Present";
      }
    }
    const ot = db.overtimes.find((o) => o.empId === emp.id && o.date === d.date);
    if (ot) day.overtime = ot.hours;
    return day;
  });
}

/* ---------- Driver and conductor batch ---------- */

export function fmsBatch(db: Db, emp: Employee, cutoff: Cutoff): FmsDay[] {
  const s = parseISO(cutoff.start);
  const e = parseISO(cutoff.end);
  const dim = daysInMonth(s.y, s.m);
  const lastDay = e.d >= dim ? 31 : e.d;
  const rows: FmsDay[] = [];
  for (let day = s.d; day <= lastDay; day++) {
    const valid = s.m === 2 ? day <= 28 : day <= Math.min(30, dim) || (dim === 31 && day === 31);
    const date = valid ? isoOf(s.y, s.m, day) : null;
    const row: FmsDay = {
      day,
      date,
      pairs: [],
      pairHours: [],
      required: 0,
      overtime: 0,
      nd: null,
      late: null,
      leaveLetter: "",
      leaveHours: 0,
      restDayOt: 0,
      remark: "Invalid Day",
    };
    if (date) {
      const ps = db.punches
        .filter((p) => p.dtrId === emp.dtrId && p.date === date)
        .sort((a, b) => a.time.localeCompare(b.time));
      let open: string | null = null;
      let unmatched = false;
      for (const p of ps) {
        if (p.dir === "In") {
          if (open) unmatched = true;
          open = p.time;
        } else if (open) {
          if (row.pairs.length < 6) row.pairs.push({ in: open, out: p.time });
          open = null;
        } else unmatched = true;
      }
      if (open) unmatched = true;
      row.pairHours = row.pairs.map((p) => {
        let b = toMin(p.out);
        const a = toMin(p.in);
        if (b < a) b += 1440;
        return r2((b - a) / 60);
      });
      row.required = r2(row.pairHours.reduce((x, y) => x + y, 0));
      row.remark = row.pairs.length
        ? row.pairHours.some((h) => h === 0)
          ? "No In/Out"
          : "Present"
        : unmatched
          ? "No In/Out"
          : "Absent";
    }
    rows.push(row);
  }
  return applyFmsLeave(db, emp, rows);
}

export function applyFmsLeave(db: Db, emp: Employee, rows: FmsDay[]): FmsDay[] {
  const lastValid = [...rows].reverse().find((r) => r.date)?.day;
  return rows.map((r) => {
    const row = { ...r, leaveLetter: "" as FmsDay["leaveLetter"], leaveHours: 0, restDayOt: 0 };
    if (!r.date) return row;
    const leave = db.leaves.find((l) => l.empId === emp.id && l.date === r.date);
    if (leave) {
      row.leaveLetter = leave.letter;
      row.leaveHours = leave.hours;
      // Rest-day overtime only on the last day, and only for letter D.
      if (leave.letter === "D" && r.day === lastValid) row.restDayOt = leave.hours;
    }
    const hol = db.holidays.find((h) => h.date === r.date);
    if (hol && !leave) {
      row.leaveLetter = hol.letter;
      row.leaveHours = r.required;
    }
    return row;
  });
}

/* ---------- Attendance totals ---------- */

const sumBy = <T>(xs: T[], f: (x: T) => number) => r2(xs.reduce((a, x) => a + f(x), 0));

export function rateBook(emp: Employee, client: Client) {
  const wd = client.workingDays ?? 303;
  const daily = emp.basis === "Monthly" ? r2((emp.monthlyRate * 12) / wd) : emp.dailyRate;
  const hourly = emp.basis === "Hourly" && emp.hourlyRate ? emp.hourlyRate : r2(daily / 8);
  return {
    daily,
    hourly,
    monthly: emp.basis === "Monthly" ? emp.monthlyRate : r2((daily * wd) / 12),
  };
}

export function attendanceTotals(
  db: Db,
  emp: Employee,
  client: Client,
  key: string,
): AttendanceTotals | null {
  const set = db.payrollAttendance[`${emp.id}|${key}`];
  if (!set) return null;
  const { daily, hourly } = rateBook(emp, client);
  if (set.kind === "staff" && set.staff) {
    const rows = set.staff;
    const letter = (ls: string[]) =>
      sumBy(
        rows.filter((r) => ls.includes(r.leaveLetter)),
        (r) => r.leaveHours,
      );
    const restHours = letter(["R"]);
    const absentDays = rows.filter((r) => r.remark === "Absent" && !r.leaveLetter).length;
    const lateHours = sumBy(rows, (r) => r.late);
    return {
      kind: "staff",
      hoursWorked: emp.basis === "Monthly" ? 104 : r2(sumBy(rows, (r) => r.required) - restHours),
      days: rows.filter((r) => r.required !== 0).length,
      overtime: sumBy(rows, (r) => r.overtime),
      nd: sumBy(rows, (r) => r.nd),
      restDay: r2(restHours + letter(["D"])),
      special: letter(["H"]),
      legal: letter(["L"]),
      sickDays: r2(letter(["M", "S"]) / 8),
      vacationDays: r2(letter(["P", "V", "Y"]) / 8),
      absentDays,
      lateHours,
      absenceAmount: emp.basis === "Monthly" ? r2(absentDays * daily) : 0,
      lateAmount: r2(lateHours * hourly),
    };
  }
  const rows = set.fms ?? [];
  const letter = (ls: string[]) =>
    sumBy(
      rows.filter((r) => ls.includes(r.leaveLetter)),
      (r) => r.leaveHours,
    );
  return {
    kind: "fms",
    hoursWorked: sumBy(rows, (r) => r.required),
    days: rows.filter((r) => r.required !== 0).length,
    overtime: 0,
    nd: 0,
    restDay: sumBy(rows, (r) => r.restDayOt),
    special: letter(["H"]),
    legal: letter(["L"]),
    sickDays: r2(letter(["M", "S"]) / 8),
    vacationDays: r2(letter(["P", "V", "Y"]) / 8),
    absentDays: rows.filter((r) => r.remark === "Absent" && !r.leaveLetter).length,
    lateHours: 0,
    absenceAmount: 0,
    lateAmount: 0,
  };
}

/* ---------- Contributions and tax ---------- */

export function sssShare(db: Db, wage: number) {
  const row = db.sss.find((r) => wage >= r.from && wage <= r.to) ?? db.sss[db.sss.length - 1];
  return row ?? { ee: 0, er: 0, ec: 0 };
}
export function philhealthShare(db: Db, wage: number) {
  const row =
    db.philhealth.find((r) => wage >= r.from && wage <= r.to) ??
    db.philhealth[db.philhealth.length - 1];
  return row ?? { ee: 0, er: 0 };
}
export function hdmfShare(db: Db, wage: number) {
  const row = db.hdmf.find((r) => wage >= r.from && wage <= r.to) ?? db.hdmf[db.hdmf.length - 1];
  const base = Math.min(wage, 10000);
  return row
    ? { ee: r2((base * row.eeRate) / 100), er: r2((base * row.erRate) / 100) }
    : { ee: 0, er: 0 };
}
export function taxDue(db: Db, table: "Semi-monthly" | "2018", code: string, taxable: number) {
  const rows = db.tax
    .filter((t) => t.table === table && (table === "2018" || t.taxCode === code))
    .sort((a, b) => a.floor - b.floor);
  const row = [...rows].reverse().find((r) => taxable >= r.floor);
  if (!row) return 0;
  return r2(row.baseTax + ((taxable - row.floor) * row.percent) / 100);
}

/* ---------- Payroll ---------- */

export function computePayroll(
  db: Db,
  emp: Employee,
  client: Client,
  cutoff: Cutoff,
  att: AttendanceTotals,
  inputs: PayInputs,
  processedBy: string,
): { record: PayrollRecord; problems: string[] } {
  const fms = isFmsClient(client);
  const key = cutoffKey(cutoff.end);
  const { daily, hourly, monthly } = rateBook(emp, client);
  const R = client.rates;
  const h = { ...att, ...inputs.hours };
  const problems: string[] = [];

  const line = (
    k: string,
    label: string,
    qty: number | null,
    unit: PayLine["unit"],
    basis: string,
    amount: number,
  ): PayLine => ({
    key: k,
    label,
    qty,
    unit,
    basis,
    amount: r2(amount),
  });
  const fmt = (n: number) =>
    n.toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const earnings: PayLine[] = [
    emp.basis === "Monthly"
      ? line("basic", "Basic pay", h.hoursWorked, "hrs", "Monthly rate / 2", monthly / 2)
      : line(
          "basic",
          "Basic pay",
          h.hoursWorked,
          "hrs",
          `${fmt(hourly)} per hour`,
          h.hoursWorked * hourly,
        ),
    line(
      "overtime",
      "Overtime",
      h.overtime,
      "hrs",
      `× ${R.overtime.toFixed(2)}`,
      h.overtime * hourly * R.overtime,
    ),
    line(
      "nd",
      "Night differential",
      h.nd,
      "hrs",
      `× ${R.nightDiff.toFixed(2)}`,
      h.nd * hourly * R.nightDiff,
    ),
    line(
      "restDay",
      "Rest day",
      h.restDay,
      "hrs",
      `× ${R.restDay.toFixed(2)}`,
      h.restDay * hourly * R.restDay,
    ),
    line(
      "special",
      "Special holiday",
      h.special,
      "hrs",
      `× ${R.specialHoliday.toFixed(2)}`,
      h.special * hourly * R.specialHoliday,
    ),
    line(
      "legal",
      "Legal holiday",
      h.legal,
      "hrs",
      `× ${R.legalHoliday.toFixed(2)}`,
      h.legal * hourly * R.legalHoliday,
    ),
    line("ecola", "ECOLA", att.days, "days", `${fmt(R.ecola)} per day`, att.days * R.ecola),
    line("hazard", "Hazard", att.days, "days", `${fmt(R.hazard)} per day`, att.days * R.hazard),
    line("allowance", "Allowance", null, "", "Entered", inputs.allowance),
    line("otherIncome", "Other income", null, "", "Entered", inputs.otherIncome),
    line("taxRefund", fms ? "Bonus" : "Tax refund", null, "", "Entered", inputs.taxRefund),
  ];
  if (client.flags.thirteenthOnCutoff) {
    earnings.push(
      line(
        "thirteenth",
        "13th month",
        att.days,
        "days",
        "Daily rate / 12",
        emp.dailyRate / 12,
      ),
    );
  }
  const gross = r2(earnings.reduce((a, l) => a + l.amount, 0));

  const monthlyWage = emp.basis === "Monthly" ? emp.monthlyRate : daily * 26;
  const deductions: PayLine[] = [];
  let contributions = 0;
  if (showsSssPhilhealth(cutoff.end)) {
    const sss = client.flags.sss && emp.sss ? sssShare(db, monthlyWage).ee : 0;
    const ph =
      client.flags.philhealth && emp.philhealth
        ? emp.philhealthPerDay > 0
          ? emp.philhealthPerDay * att.days
          : philhealthShare(db, monthlyWage).ee
        : 0;
    deductions.push(line("sss", "SSS employee", null, "", "Table", sss));
    deductions.push(
      line(
        "philhealth",
        "PhilHealth employee",
        emp.philhealthPerDay > 0 ? att.days : null,
        emp.philhealthPerDay > 0 ? "days" : "",
        emp.philhealthPerDay > 0 ? `${fmt(emp.philhealthPerDay)} per day` : "Table",
        ph,
      ),
    );
    contributions += sss + ph;
  } else {
    const hd =
      client.flags.pagibig && emp.pagibig
        ? emp.pagibigPerDay > 0
          ? emp.pagibigPerDay * att.days
          : hdmfShare(db, monthlyWage).ee
        : 0;
    deductions.push(
      line(
        "pagibig",
        "Pag-IBIG employee",
        emp.pagibigPerDay > 0 ? att.days : null,
        emp.pagibigPerDay > 0 ? "days" : "",
        emp.pagibigPerDay > 0 ? `${fmt(emp.pagibigPerDay)} per day` : "Table",
        hd,
      ),
    );
    contributions += hd;
  }

  let tax = 0;
  const table = client.flags.tax2018 ? "2018" : "Semi-monthly";
  if (!emp.taxCode)
    problems.push("No tax code on this employee. Add a tax code in Employees before saving.");
  else if (
    table === "Semi-monthly" &&
    !db.tax.some((t) => t.table === "Semi-monthly" && t.taxCode === emp.taxCode)
  )
    problems.push(
      `Tax code ${emp.taxCode} is missing from the semi-monthly tax table. Add it in Maintenance before saving.`,
    );
  else if (!(client.taxPercent && daily <= client.taxCeiling))
    tax = taxDue(
      db,
      table,
      emp.taxCode,
      gross - contributions - att.absenceAmount - att.lateAmount,
    );
  deductions.push(
    line(
      "tax",
      "Withholding tax",
      null,
      "",
      client.taxPercent && daily <= client.taxCeiling
        ? `Zero: daily rate at or under ${fmt(client.taxCeiling)}`
        : `${table} table, ${emp.taxCode || "no code"}`,
      tax,
    ),
  );

  if (showsUnionDues(cutoff.end))
    deductions.push(line("union", "Union dues", null, "", "First half only", fms ? 25 : 50));

  const half = cutoffHalf(cutoff.end);
  for (const d of db.deductions.filter((x) => x.empId === emp.id && x.balance > 0)) {
    if (d.schedule === "First half" && half !== 1) continue;
    if (d.schedule === "Second half" && half !== 2) continue;
    deductions.push(
      line(
        `ded-${d.code}-${d.id}`,
        deductionLabel(d.code, fms),
        null,
        "",
        `Balance ${fmt(d.balance)}`,
        Math.min(d.amount, d.balance),
      ),
    );
  }
  deductions.push(
    line(
      "absence",
      "Absence",
      att.absentDays,
      "days",
      emp.basis === "Monthly" ? `${fmt(daily)} per day` : "Not deducted on hourly basis",
      att.absenceAmount,
    ),
  );
  deductions.push(
    line("late", "Late", att.lateHours, "hrs", `${fmt(hourly)} per hour`, att.lateAmount),
  );

  const totalDeductions = r2(deductions.reduce((a, l) => a + l.amount, 0));
  return {
    problems,
    record: {
      id: uid(),
      empId: emp.id,
      clientId: client.id,
      key,
      start: cutoff.start,
      end: cutoff.end,
      half,
      attendance: { ...att, ...inputs.hours },
      earnings,
      deductions,
      gross,
      totalDeductions,
      net: r2(gross - totalDeductions),
      inputs,
      rates: { basis: emp.basis, monthly, daily, hourly, taxCode: emp.taxCode },
      processedBy,
      savedAt: new Date().toISOString(),
      movedAt: null,
    },
  };
}

/* ---------- Driver commissions ---------- */

export function commissionContributions(db: Db, commission: number, ceiling: number) {
  const capped = Math.min(commission, ceiling);
  // Driver SSS and PhilHealth keep the old minwage >= commission test. Pag-IBIG stays BETWEEN.
  const sss = db.sss.find((r) => r.from >= capped && capped <= r.to);
  const ph = db.philhealth.find((r) => r.from >= capped && capped <= r.to);
  const hd = db.hdmf.find((r) => capped >= r.from && capped <= r.to);
  return {
    capped,
    sssEe: sss?.ee ?? 0,
    sssEr: sss?.er ?? 0,
    ecc: sss?.ec ?? 0,
    phEe: ph?.ee ?? 0,
    phEr: ph?.er ?? 0,
    hdmfEe: hd ? r2((capped * hd.eeRate) / 100) : 0,
    hdmfEr: hd ? r2((capped * hd.erRate) / 100) : 0,
  };
}
