import { useState } from "react";
import { useSession } from "../app/session";
import { ApiError, generateThirteenth } from "../data/api";
import type { Thirteenth as T13 } from "../data/types";
import { fullName } from "../data/rules";
import {
  Button,
  Field,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  StatusLine,
  TextInput,
  useStatus,
  EmptyState,
} from "../ui/kit";
import { Icon } from "../ui/icons";
import { fmtDate, money } from "../ui/format";

export function Thirteenth() {
  const { db, clients } = useSession();
  const [clientId, setClientId] = useState(clients[0]?.id ?? "");
  const [year, setYear] = useState(2025);
  const people = db.employees
    .filter((e) => e.clientId === clientId)
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const [empId, setEmpId] = useState(people[0]?.id ?? "");
  const [from, setFrom] = useState("2025-01-01");
  const [to, setTo] = useState("2025-12-15");
  const [row, setRow] = useState<T13 | null>(null);
  const [busy, setBusy] = useState(false);
  const [field, setField] = useState<string | undefined>();
  const { status, setStatus, clear } = useStatus();
  const emp = db.employees.find((e) => e.id === empId);
  const client = db.clients.find((c) => c.id === clientId);

  const reset = () => {
    setRow(null);
    clear();
  };
  const gen = async () => {
    setBusy(true);
    setField(undefined);
    try {
      const r = await generateThirteenth(empId, year);
      setRow(r);
      setStatus({
        tone: "success",
        title: `Generated the ${year} 13th-month payslip for ${emp ? `${emp.empNo} ${fullName(emp)}` : ""}.`,
      });
    } catch (e) {
      setRow(null);
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({
        tone: "error",
        title: "13th-month payslip not generated",
        body: (e as Error).message,
      });
    } finally {
      setBusy(false);
    }
  };

  return (
    <>
      <PageHeader
        section="Payroll"
        title="13th-month payslip"
        description="Year-end print from the stored 13th-month amount. It does not divide the rate again."
      />
      <div className="grid items-start gap-4 xl:grid-cols-[380px_minmax(0,1fr)]">
        <Panel className="no-print">
          <PanelHeader
            title="Printout"
            actions={
              <>
                <Button icon={<Icon.printer />} onClick={() => window.print()} disabled={!row}>
                  Print
                </Button>
                <Button variant="primary" onClick={gen} busy={busy}>
                  Generate
                </Button>
              </>
            }
          />
          <div className="space-y-3 px-4 py-4">
            <StatusLine status={status} onDismiss={clear} />
            <Field label="Client">
              <Select
                value={clientId}
                onChange={(e) => {
                  setClientId(e.target.value);
                  setEmpId(db.employees.find((x) => x.clientId === e.target.value)?.id ?? "");
                  reset();
                }}
              >
                {clients.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Year" error={field === "year" ? (status?.body as string) : null}>
              <Select value={year} onChange={(e) => (setYear(Number(e.target.value)), reset())}>
                {[2025, 2024, 2023].map((y) => (
                  <option key={y}>{y}</option>
                ))}
              </Select>
            </Field>
            <Field label="Employee" error={field === "empId" ? (status?.body as string) : null}>
              <Select value={empId} onChange={(e) => (setEmpId(e.target.value), reset())}>
                {people.map((e) => (
                  <option key={e.id} value={e.id}>
                    {e.empNo} · {fullName(e)}
                  </option>
                ))}
              </Select>
            </Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Pay period from">
                <TextInput type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
              </Field>
              <Field label="Pay period to">
                <TextInput type="date" value={to} onChange={(e) => setTo(e.target.value)} />
              </Field>
            </div>
            <p className="text-xs text-grey-500">
              The period dates print on the payslip. They do not change the amount.
            </p>
          </div>
        </Panel>

        {row && emp && client ? (
          <article className="print-flat w-full max-w-[640px] rounded-panel border border-grey-100 bg-white text-sm">
            <header className="flex items-start justify-between gap-6 border-b border-grey-200 px-6 py-4">
              <div>
                {db.company?.name ? (
                  <>
                    <div className="text-base font-semibold">{db.company.name}</div>
                    <div className="text-xs text-grey-600">{db.company.address}</div>
                  </>
                ) : (
                  <div className="no-print text-xs text-grey-500">
                    Report header is blank until the company is saved.
                  </div>
                )}
              </div>
              <div className="text-right">
                <div className="text-base font-semibold">13th-month pay</div>
                <div className="text-xs text-grey-600">Year {row.year}</div>
              </div>
            </header>
            <div className="grid grid-cols-2 gap-x-6 gap-y-2 border-b border-grey-100 px-6 py-3">
              <Meta label="Employee" value={`${emp.empNo} ${fullName(emp)}`} />
              <Meta label="Client" value={client.name} />
              <Meta label="Position" value={emp.position} />
              <Meta label="Pay period" value={`${fmtDate(from)} to ${fmtDate(to)}`} />
            </div>
            <table className="w-full">
              <tbody>
                <tr className="border-b border-grey-100">
                  <td className="h-11 px-6 text-grey-700">Gross 13th-month pay</td>
                  <td className="px-6 text-right font-medium">{money(row.amount)}</td>
                </tr>
                <tr>
                  <td className="h-12 px-6 font-semibold text-grey-900">Take-home</td>
                  <td className="px-6 text-right text-xl font-semibold">{money(row.amount)}</td>
                </tr>
              </tbody>
            </table>
            <footer className="border-t border-grey-100 px-6 py-3 text-xs text-grey-500">
              Stored year amount. Printed as saved.
            </footer>
          </article>
        ) : (
          <Panel className="no-print">
            <EmptyState>Choose the client, year, and employee, then Generate.</EmptyState>
          </Panel>
        )}
      </div>
    </>
  );
}

function Meta({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <div className="text-xs text-grey-500">{label}</div>
      <div className="font-medium text-grey-900">{value}</div>
    </div>
  );
}
