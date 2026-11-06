import type { Company, Client, Employee, PayLine, PayrollRecord } from "../data/types";
import { fullName, isFmsClient } from "../data/rules";
import { fmtDate, fmtRange, money } from "../ui/format";
import { cx } from "../ui/kit";

type Row = { label: string; amount: number };

/** Map saved lines to payslip labels. Staff and FMS print different named rows. */
function payslipRows(record: PayrollRecord, fms: boolean) {
  const earnings: Row[] = record.earnings
    .filter((l) => l.amount !== 0 || ["basic", "taxRefund"].includes(l.key))
    .map((l) => ({ label: l.label, amount: l.amount }));
  const byCode = (l: PayLine) => (l.key.startsWith("ded-") ? Number(l.key.split("-")[1]) : null);
  const deductions: Row[] = [];
  const add = (label: string, amount: number) => {
    const existing = deductions.find((d) => d.label === label);
    if (existing) existing.amount += amount;
    else deductions.push({ label, amount });
  };
  for (const l of record.deductions) {
    const code = byCode(l);
    if (fms && code === 4) add("Damages", l.amount);
    else if (fms && code === 8) add("Shortages", l.amount);
    else if (fms && (code === 16 || code === 17)) add("Calamity loan", l.amount);
    else if (!fms && code === 8) add("Other deductions", l.amount);
    else if (l.amount !== 0 || ["sss", "philhealth", "pagibig", "tax"].includes(l.key))
      add(l.label, l.amount);
  }
  const required = fms
    ? ["Damages", "Shortages", "Calamity loan"]
    : ["Planters", "Other deductions"];
  for (const r of required)
    if (!deductions.some((d) => d.label === r)) deductions.push({ label: r, amount: 0 });
  return { earnings, deductions };
}

export function PayslipDoc({
  record,
  emp,
  client,
  company,
  source,
}: {
  record: PayrollRecord;
  emp: Employee;
  client: Client;
  company: Company | null;
  source: "open" | "history";
}) {
  const fms = isFmsClient(client);
  const { earnings, deductions } = payslipRows(record, fms);
  const days = record.attendance.hoursWorked / 8;
  return (
    <article className="print-flat mx-auto w-full max-w-[720px] rounded-panel border border-grey-100 bg-white text-sm">
      <header className="flex items-start justify-between gap-6 border-b border-grey-200 px-6 py-4">
        <div>
          {company?.name ? (
            <>
              <div className="text-base font-semibold text-grey-900">{company.name}</div>
              <div className="text-xs text-grey-600">{company.address}</div>
            </>
          ) : (
            <div className="no-print text-xs text-grey-500">
              Report header is blank until the company is saved.
            </div>
          )}
        </div>
        <div className="text-right">
          <div className="text-base font-semibold text-grey-900">Payslip</div>
          <div className="text-xs text-grey-600">
            {source === "history" ? "From payroll history" : "Open payroll, saved"}
          </div>
        </div>
      </header>
      <div className="grid grid-cols-2 gap-x-6 gap-y-1.5 border-b border-grey-100 px-6 py-3 sm:grid-cols-4">
        <Meta label="Employee" value={fullName(emp)} />
        <Meta label="Employee number" value={emp.empNo} />
        <Meta label="Client" value={client.name} />
        <Meta label="Cutoff" value={fmtRange(record.start, record.end)} />
        <Meta label="Position" value={emp.position} />
        <Meta label="Tax code" value={record.rates.taxCode || "None"} />
        <Meta label="Days" value={days.toFixed(2)} />
        <Meta label="DTR key" value={record.key} />
      </div>
      <div className="grid sm:grid-cols-2">
        <Column title="Earnings" rows={earnings} total={record.gross} totalLabel="Gross" />
        <Column
          title="Deductions"
          rows={deductions}
          total={record.totalDeductions}
          totalLabel="Total deductions"
          className="border-t border-grey-100 sm:border-l sm:border-t-0"
        />
      </div>
      <footer className="flex items-center justify-between border-t-2 border-grey-800 px-6 py-3">
        <span className="text-xs text-grey-600">
          Processed by {record.processedBy} · {fmtDate(record.savedAt)}
        </span>
        <div className="text-right">
          <div className="text-xs text-grey-500">Net pay</div>
          <div className="text-xl font-semibold text-grey-900">{money(record.net)}</div>
        </div>
      </footer>
    </article>
  );
}

function Meta({ label, value }: { label: string; value: string }) {
  return (
    <div className="min-w-0">
      <div className="text-xs text-grey-500">{label}</div>
      <div className="truncate font-medium text-grey-900">{value}</div>
    </div>
  );
}

function Column({
  title,
  rows,
  total,
  totalLabel,
  className,
}: {
  title: string;
  rows: Row[];
  total: number;
  totalLabel: string;
  className?: string;
}) {
  return (
    <div className={cx("px-6 py-3", className)}>
      <div className="mb-1.5 text-xs font-semibold text-grey-600">{title}</div>
      <table className="w-full">
        <tbody>
          {rows.map((r) => (
            <tr key={r.label}>
              <td className="py-1 text-grey-800">{r.label}</td>
              <td
                className={cx(
                  "py-1 text-right",
                  r.amount === 0 ? "text-grey-400" : "text-grey-900",
                )}
              >
                {money(r.amount)}
              </td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr className="border-t border-grey-200">
            <td className="pt-2 font-semibold text-grey-900">{totalLabel}</td>
            <td className="pt-2 text-right font-semibold text-grey-900">{money(total)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
  );
}
