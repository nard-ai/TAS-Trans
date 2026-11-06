import { useEffect, useState, type ReactNode } from "react";
import { useSession } from "../app/session";
import { navigate } from "../app/router";
import { loadAttendance, moveToMaster, recompute, savePayroll } from "../data/api";
import type {
  AttendanceTotals,
  Employee,
  HourKey,
  PayInputs,
  PayLine,
  PayrollRecord,
} from "../data/types";
import { fullName, isExcluded, showsSssPhilhealth, showsUnionDues } from "../data/rules";
import { emptyInputs } from "../data/money";
import {
  Amount,
  Button,
  DataTable,
  EmptyState,
  KeyValues,
  Notice,
  NumberInput,
  PageHeader,
  Panel,
  PanelHeader,
  StatusLine,
  Tag,
  cx,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { Icon } from "../ui/icons";
import { fmtDate, fmtDateTime, hrs, money } from "../ui/format";

type Draft = { record: PayrollRecord; problems: string[]; inputs: PayInputs };

export function PayrollEntry({ param }: { param?: string }) {
  const s = useSession();
  const { db, user, cutoff, cutoffClient, cutoffKey, cutoffHalf, billed, fms } = s;
  const confirm = useConfirm();
  const page = useStatus();
  const [moving, setMoving] = useState(false);
  const [empId, setEmpId] = useState<string | null>(param ?? null);

  if (!cutoff || !cutoffClient || !cutoffKey) {
    return (
      <>
        <PageHeader section="Payroll" title="Payroll entry" />
        <Notice
          tone="error"
          title="No cutoff selected"
          actions={<Button onClick={() => navigate("cutoff")}>Select cutoff</Button>}
        >
          Payroll entry needs a client and period. The employee list stays closed until a cutoff is
          selected.
        </Notice>
      </>
    );
  }

  const people = db.employees
    .filter((e) => e.clientId === cutoffClient.id && !isExcluded(e))
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const source = billed ? db.history : db.openPayroll;
  const savedFor = (id: string) =>
    source.find((r) => r.empId === id && r.key === cutoffKey && r.clientId === cutoffClient.id);
  const processed = people.filter((e) => savedFor(e.id)).length;
  const openRows = db.openPayroll.filter(
    (r) => r.key === cutoffKey && r.clientId === cutoffClient.id,
  ).length;

  const onMove = async () => {
    const ok = await confirm({
      title: "Move to master",
      tone: "danger",
      confirmLabel: "Move to master",
      acknowledge: "I understand this closes the cutoff for the whole client.",
      body: (
        <div className="space-y-2">
          <p>
            This closes <b>{cutoffClient.name}</b>, cutoff <b>{cutoffKey}</b>, for the whole client.
            It copies the open payroll into history, removes the open rows, and marks the cutoff
            billed. After this, the cutoff is read-only.
          </p>
          <p>
            <b>{openRows}</b> saved {openRows === 1 ? "employee moves" : "employees move"} to
            history.
            {people.length - processed > 0 && (
              <span className="font-medium text-red-600">
                {" "}
                {people.length - processed} unprocessed{" "}
                {people.length - processed === 1 ? "employee is" : "employees are"} not included.
              </span>
            )}
          </p>
        </div>
      ),
    });
    if (!ok) return;
    setMoving(true);
    try {
      const r = await moveToMaster();
      page.setStatus({
        tone: "success",
        title: `Moved ${r.count} payroll rows for ${cutoffClient.name}, cutoff ${cutoffKey}, to history.`,
        body: "The cutoff is billed and read-only. Payslips print from history.",
      });
    } catch (e) {
      page.setStatus({ tone: "error", title: "Move to master failed", body: (e as Error).message });
    } finally {
      setMoving(false);
    }
  };

  return (
    <>
      <PageHeader
        section="Payroll"
        title="Payroll entry"
        description="One employee, one cutoff. Load attendance, check what the server returns, then save."
        actions={
          !billed && (
            <Button variant="danger" onClick={onMove} busy={moving} disabled={openRows === 0}>
              Move to master
            </Button>
          )
        }
      />

      <Panel className="mb-4">
        <div className="px-4 py-3">
          <KeyValues
            cols={4}
            items={[
              { label: "Client", value: cutoffClient.name },
              { label: "Cutoff", value: `${fmtDate(cutoff.start)} to ${fmtDate(cutoff.end)}` },
              { label: "Half and DTR key", value: `Half ${cutoffHalf} · ${cutoffKey}` },
              { label: "Type and user", value: `${fms ? "FMS" : "Staff"} · ${user.username}` },
            ]}
          />
        </div>
      </Panel>

      {(page.status || billed) && (
        <div className="mb-4 space-y-2">
          <StatusLine status={page.status} onDismiss={page.clear} />
          {billed && (
            <Notice tone="info" title="This cutoff is billed and read-only">
              Payroll was moved to history. Every field is read-only. Print payslip still works.
            </Notice>
          )}
        </div>
      )}

      <div className="grid items-start gap-4 xl:grid-cols-[260px_minmax(0,1fr)]">
        <Panel className={cx(empId && "hidden xl:block")}>
          <PanelHeader
            title="Employees"
            meta={`${processed} processed · ${people.length - processed} remaining`}
          />
          <DataTable
            rows={people}
            rowKey={(e) => e.id}
            selected={empId}
            onSelect={(e) => setEmpId(e.id)}
            maxHeight="calc(100vh - 330px)"
            empty={<EmptyState>No employees on this client are included in payroll.</EmptyState>}
            columns={[
              {
                key: "n",
                header: "Employee",
                render: (e) => (
                  <div className="leading-tight">
                    <div className="font-medium text-grey-900">{fullName(e)}</div>
                    <div className="text-xs text-grey-500">{e.empNo}</div>
                  </div>
                ),
              },
              {
                key: "p",
                header: "Payroll",
                align: "right",
                render: (e) =>
                  savedFor(e.id) ? (
                    <span className="text-xs font-medium text-cyan-800">Processed</span>
                  ) : (
                    <span className="text-xs text-grey-400">Not processed</span>
                  ),
              },
            ]}
          />
        </Panel>

        {empId && db.employees.find((e) => e.id === empId) ? (
          <EmployeePay
            key={empId}
            emp={db.employees.find((e) => e.id === empId)!}
            saved={savedFor(empId)}
            onBack={() => setEmpId(null)}
          />
        ) : (
          <Panel>
            <PanelHeader title="No employee open" />
            <EmptyState>
              Choose an employee on the left. Processed employees show their saved payroll.
            </EmptyState>
          </Panel>
        )}
      </div>
    </>
  );
}

const HOUR_ROWS: { key: HourKey; label: string }[] = [
  { key: "hoursWorked", label: "Hours worked" },
  { key: "overtime", label: "Regular overtime" },
  { key: "nd", label: "Night differential" },
  { key: "restDay", label: "Rest day" },
  { key: "special", label: "Special holiday (H)" },
  { key: "legal", label: "Legal holiday (L)" },
];

function EmployeePay({
  emp,
  saved,
  onBack,
}: {
  emp: Employee;
  saved: PayrollRecord | undefined;
  onBack: () => void;
}) {
  const { cutoff, cutoffKey, billed, fms, db } = useSession();
  const { status, setStatus, clear } = useStatus();
  const [draft, setDraft] = useState<Draft | null>(null);
  const [inputs, setInputs] = useState<PayInputs>(saved?.inputs ?? emptyInputs());
  const [busy, setBusy] = useState<"load" | "recompute" | "save" | null>(null);
  const client = db.clients.find((c) => c.id === emp.clientId)!;
  const who = `${emp.empNo} ${fullName(emp)}`;

  useEffect(() => {
    if (!draft) setInputs(saved?.inputs ?? emptyInputs());
  }, [saved, draft]);

  const record = draft?.record ?? saved ?? null;
  const editing = !!draft && !billed;
  const stale = !!draft && JSON.stringify(inputs) !== JSON.stringify(draft.inputs);
  const att: AttendanceTotals | null = record ? record.attendance : null;
  const rates = record?.rates;

  const doLoad = async () => {
    setBusy("load");
    clear();
    try {
      const base = {
        ...emptyInputs(),
        allowance: inputs.allowance,
        otherIncome: inputs.otherIncome,
        taxRefund: inputs.taxRefund,
      };
      const r = await loadAttendance(emp.id, base);
      setDraft({ record: r.record, problems: r.problems, inputs: base });
      setInputs(base);
      setStatus({
        tone: "info",
        title: `Loaded finalized attendance for ${who}.`,
        body: "Amounts are from the server. Change hours if needed, then Save.",
      });
    } catch (e) {
      setStatus({ tone: "error", title: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const doRecompute = async () => {
    setBusy("recompute");
    try {
      const r = await recompute(emp.id, inputs);
      setDraft({ record: r.record, problems: r.problems, inputs });
      setStatus({ tone: "info", title: "The server refilled the amounts." });
    } catch (e) {
      setStatus({ tone: "error", title: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const doSave = async () => {
    setBusy("save");
    try {
      await savePayroll(emp.id, inputs);
      setDraft(null);
      setStatus({
        tone: "success",
        title: `Saved payroll for ${who}, cutoff ${cutoffKey}.`,
        body: "Marked processed. Print payslip is ready.",
      });
    } catch (e) {
      setStatus({ tone: "error", title: "Payroll not saved", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };

  const setHour = (k: HourKey, v: number | null) =>
    setInputs((i) => ({ ...i, hours: { ...i.hours, [k]: v ?? 0 } }));
  const setEntered = (k: "allowance" | "otherIncome" | "taxRefund", v: number | null) =>
    setInputs((i) => ({ ...i, [k]: v ?? 0 }));
  const hourValue = (k: HourKey) => inputs.hours[k] ?? (draft?.record.attendance[k] as number) ?? 0;

  const half = cutoff
    ? showsSssPhilhealth(cutoff.end)
      ? "SSS and PhilHealth shown; Pag-IBIG hidden this half."
      : "Pag-IBIG shown; SSS and PhilHealth hidden this half."
    : "";
  const union = cutoff && !showsUnionDues(cutoff.end) ? " Union dues are zero after the 15th." : "";
  const loadPrimary = !billed && !draft && !saved;

  return (
    <Panel className="min-w-0">
      <PanelHeader
        title={
          <span className="flex items-center gap-2">
            <button
              type="button"
              onClick={onBack}
              className="-ml-1 rounded-ctl p-1 text-grey-500 hover:bg-grey-50 hover:text-grey-800 xl:hidden"
              aria-label="Back to employee list"
            >
              <Icon.arrowLeft />
            </button>
            {who}
          </span>
        }
        meta={
          saved && !draft
            ? `${billed ? "In history" : "Processed"} · saved ${fmtDateTime(saved.savedAt)} by ${saved.processedBy}`
            : draft
              ? stale
                ? "Hours changed, amounts out of date"
                : "Loaded, not saved"
              : "Not processed"
        }
        actions={
          <>
            <Button variant="quiet" onClick={onBack} icon={<Icon.arrowLeft />}>
              Back to employee list
            </Button>
            <Button
              onClick={() => navigate(`payslip/${emp.id}`)}
              disabled={!saved || !!draft}
              icon={<Icon.printer />}
            >
              Print payslip
            </Button>
            {!billed && (
              <>
                <Button
                  variant={loadPrimary ? "primary" : "secondary"}
                  onClick={doLoad}
                  busy={busy === "load"}
                  disabled={!!busy}
                >
                  Load attendance
                </Button>
                <Button
                  onClick={doRecompute}
                  busy={busy === "recompute"}
                  disabled={!draft || !!busy}
                >
                  Recompute
                </Button>
                <Button
                  variant={loadPrimary ? "secondary" : "primary"}
                  onClick={doSave}
                  busy={busy === "save"}
                  disabled={!draft || !!busy}
                >
                  Save
                </Button>
              </>
            )}
          </>
        }
      />

      <div className="flex flex-wrap gap-x-8 gap-y-2 border-b border-grey-100 bg-grey-50/60 px-4 py-2.5 text-sm">
        <Fact label="Pay basis" value={emp.basis} />
        <Fact
          label="Monthly"
          value={money(rates?.monthly ?? (emp.basis === "Monthly" ? emp.monthlyRate : null))}
        />
        <Fact
          label="Daily"
          value={money(rates?.daily ?? (emp.basis === "Hourly" ? emp.dailyRate : null))}
        />
        <Fact
          label="Hourly"
          value={money(
            rates?.hourly ?? (emp.hourlyRate || (emp.dailyRate ? emp.dailyRate / 8 : null)),
          )}
        />
        <Fact
          label="Tax code"
          value={emp.taxCode || <span className="font-medium text-red-600">None</span>}
        />
        <Fact label="Type" value={fms ? "FMS driver attendance" : "Staff attendance"} />
      </div>

      {(status || (draft && draft.problems.length > 0)) && (
        <div className="space-y-2 border-b border-grey-100 px-4 py-3">
          <StatusLine status={status} onDismiss={clear} />
          {draft?.problems.map((p) => (
            <Notice key={p} tone="error" title="Save will fail until this is fixed">
              {p}
            </Notice>
          ))}
        </div>
      )}

      {!record || !att ? (
        <EmptyState
          action={
            !billed && <span className="text-xs text-grey-500">Use Load attendance above.</span>
          }
        >
          {billed
            ? "No payroll was saved for this employee before the cutoff was billed."
            : "Attendance is not loaded. Load attendance pulls the finalized time records for this cutoff."}
        </EmptyState>
      ) : (
        <>
          <div className={cx("grid gap-0 lg:grid-cols-[300px_minmax(0,1fr)]", stale && "relative")}>
            <div className="border-b border-grey-100 lg:border-b-0 lg:border-r">
              <SectionTitle
                title="Attendance"
                note={editing ? "Hours are editable until saved." : undefined}
              />
              <table className="w-full text-sm">
                <tbody>
                  {HOUR_ROWS.map((h) => (
                    <tr key={h.key} className="border-b border-grey-100">
                      <th
                        scope="row"
                        className="h-9 whitespace-nowrap px-4 text-left font-normal text-grey-700"
                      >
                        {h.label}
                        {h.key === "hoursWorked" && emp.basis === "Monthly" && (
                          <span className="ml-1 text-xs text-grey-400">monthly: 104</span>
                        )}
                      </th>
                      <td className="w-28 px-3 text-right">
                        {editing ? (
                          <NumberInput
                            aria-label={`${h.label} hours`}
                            className="h-7"
                            value={hourValue(h.key)}
                            onValue={(v) => setHour(h.key, v)}
                          />
                        ) : (
                          <span className="font-medium">{hrs(att[h.key])}</span>
                        )}
                      </td>
                      <td className="w-10 pr-4 text-xs text-grey-400">hrs</td>
                    </tr>
                  ))}
                  <AttRow label="Days with required hours" value={String(att.days)} unit="days" />
                  <AttRow label="Sick leave" value={att.sickDays.toFixed(2)} unit="days" />
                  <AttRow label="Vacation leave" value={att.vacationDays.toFixed(2)} unit="days" />
                  <AttRow
                    label="Absence"
                    value={money(att.absenceAmount)}
                    unit=""
                    sub={`${att.absentDays} days`}
                    bad={att.absenceAmount > 0}
                  />
                  <AttRow
                    label="Late"
                    value={money(att.lateAmount)}
                    unit=""
                    sub={`${hrs(att.lateHours)} hrs`}
                    bad={att.lateAmount > 0}
                  />
                </tbody>
              </table>
            </div>

            <div className={cx("min-w-0", stale && "opacity-50")}>
              <SectionTitle title="Earnings" note="Amounts come from the server." />
              <Lines
                lines={record.earnings}
                editing={editing}
                inputs={inputs}
                setEntered={setEntered}
              />
              <SectionTitle title="Deductions" note={`${half}${union}`} />
              <Lines
                lines={record.deductions}
                editing={false}
                inputs={inputs}
                setEntered={setEntered}
                negative
              />
            </div>
          </div>

          <div className="sticky bottom-0 z-[2] flex flex-wrap items-center justify-end gap-x-10 gap-y-2 rounded-b-panel border-t border-grey-200 bg-white px-4 py-3 shadow-edge">
            {stale && (
              <p className="mr-auto flex items-center gap-1.5 text-sm font-medium text-red-600">
                <Icon.alert className="size-4" />
                Hours changed. Recompute or Save to refresh amounts.
              </p>
            )}
            <Total label="Gross" value={record.gross} muted={stale} />
            <Total label="Deductions" value={record.totalDeductions} muted={stale} />
            <Total label="Net pay" value={record.net} big muted={stale} />
          </div>
        </>
      )}
      {client && fms && (
        <p className="px-4 pb-3 pt-2 text-xs text-grey-500">
          Trip commissions are not imported here. They are paid on Driver commissions.
        </p>
      )}
    </Panel>
  );
}

function Fact({ label, value }: { label: string; value: ReactNode }) {
  return (
    <div className="flex items-baseline gap-1.5">
      <span className="text-xs text-grey-500">{label}</span>
      <span className="font-medium text-grey-900">{value}</span>
    </div>
  );
}

function SectionTitle({ title, note }: { title: string; note?: string }) {
  return (
    <div className="flex items-baseline justify-between gap-3 border-b border-grey-100 bg-grey-50 px-4 py-2">
      <h3 className="text-sm font-semibold text-grey-900">{title}</h3>
      {note && <span className="max-w-[60%] text-xs text-grey-500">{note}</span>}
    </div>
  );
}

function AttRow({
  label,
  value,
  unit,
  sub,
  bad,
}: {
  label: string;
  value: string;
  unit: string;
  sub?: string;
  bad?: boolean;
}) {
  return (
    <tr className="border-b border-grey-100 last:border-b-0">
      <th scope="row" className="h-9 whitespace-nowrap px-4 text-left font-normal text-grey-700">
        {label}
        {sub && <span className="ml-1 text-xs text-grey-400">{sub}</span>}
      </th>
      <td className={cx("px-3 text-right font-medium", bad ? "text-red-600" : "text-grey-900")}>
        {value}
      </td>
      <td className="pr-4 text-xs text-grey-400">{unit}</td>
    </tr>
  );
}

function Lines({
  lines,
  editing,
  inputs,
  setEntered,
  negative,
}: {
  lines: PayLine[];
  editing: boolean;
  inputs: PayInputs;
  setEntered: (k: "allowance" | "otherIncome" | "taxRefund", v: number | null) => void;
  negative?: boolean;
}) {
  const entered = (k: string): k is "allowance" | "otherIncome" | "taxRefund" =>
    k === "allowance" || k === "otherIncome" || k === "taxRefund";
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr className="text-xs text-grey-500">
            <th className="h-7 px-4 text-left font-medium">Item</th>
            <th className="w-24 whitespace-nowrap px-3 text-right font-medium">Hours or qty</th>
            <th className="px-3 text-left font-medium">Basis</th>
            <th className="w-32 px-4 text-right font-medium">Amount</th>
          </tr>
        </thead>
        <tbody>
          {lines.map((l) => (
            <tr key={l.key} className="border-t border-grey-100">
              <td className="h-9 whitespace-nowrap px-4 text-grey-900">
                {l.label}
                {l.key === "thirteenth" && (
                  <Tag tone="blue" className="ml-2">
                    Daily rate / 12
                  </Tag>
                )}
              </td>
              <td className="whitespace-nowrap px-3 text-right text-grey-700">
                {l.qty == null ? (
                  <span className="text-grey-300">—</span>
                ) : (
                  `${l.unit === "days" ? l.qty : hrs(l.qty)} ${l.unit}`
                )}
              </td>
              <td className="whitespace-nowrap px-3 text-xs text-grey-500">{l.basis}</td>
              <td className="px-4 text-right">
                {editing && entered(l.key) ? (
                  <NumberInput
                    aria-label={l.label}
                    className="h-7"
                    value={inputs[l.key]}
                    onValue={(v) => setEntered(l.key as "allowance", v)}
                  />
                ) : (
                  <Amount
                    value={l.amount}
                    negative={negative && l.amount > 0 && (l.key === "absence" || l.key === "late")}
                  />
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Total({
  label,
  value,
  big,
  muted,
}: {
  label: string;
  value: number;
  big?: boolean;
  muted?: boolean;
}) {
  return (
    <div className="text-right">
      <div className="text-xs text-grey-500">{label}</div>
      <div
        className={cx(
          "font-semibold",
          big ? "text-xl text-grey-900" : "text-base text-grey-800",
          muted && "text-grey-400 line-through decoration-grey-300",
        )}
      >
        {money(value)}
      </div>
    </div>
  );
}
