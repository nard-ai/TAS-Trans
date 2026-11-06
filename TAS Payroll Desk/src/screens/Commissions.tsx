import { useState } from "react";
import { useSession } from "../app/session";
import { ApiError, loadTrips, saveCommissionPayment } from "../data/api";
import type { CommissionPayment, Trip } from "../data/types";
import { MONTHS, fullName, isExcluded, isFmsClient } from "../data/rules";
import { r2 } from "../data/money";
import {
  Amount,
  Button,
  DataTable,
  EmptyState,
  Field,
  Notice,
  NumberInput,
  PageHeader,
  Panel,
  PanelHeader,
  Readout,
  Select,
  StatusLine,
  TextInput,
  useStatus,
} from "../ui/kit";
import { fmtDate, fmtDateTime, money } from "../ui/format";

export function Commissions() {
  const { db, clients } = useSession();
  const fmsIds = new Set(clients.filter(isFmsClient).map((c) => c.id));
  const [role, setRole] = useState<"Driver" | "Conductor">("Driver");
  const people = db.employees
    .filter((e) => fmsIds.has(e.clientId) && !isExcluded(e) && e.role === role)
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const [empId, setEmpId] = useState(people[0]?.id ?? "");
  const emp = db.employees.find((e) => e.id === empId);
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(11);
  const [dbName, setDbName] = useState("fms_trips_2025");
  const [taxCode, setTaxCode] = useState(emp?.taxCode ?? "");
  const [trips, setTrips] = useState<Trip[] | null>(null);
  const [ceiling, setCeiling] = useState<number | null>(20000);
  const [beginning, setBeginning] = useState<number | null>(0);
  const [tax, setTax] = useState<number | null>(200);
  const [result, setResult] = useState<CommissionPayment | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [field, setField] = useState<string | undefined>();
  const [quarter, setQuarter] = useState("");
  const [listYear, setListYear] = useState(2025);
  const { status, setStatus, clear } = useStatus();

  const commission = r2((trips ?? []).reduce((a, t) => a + t.commission, 0));
  const fund = r2((trips ?? []).reduce((a, t) => a + t.fund, 0));
  const capped = ceiling != null && commission > ceiling;
  const who = emp ? `${emp.empNo} ${fullName(emp)}` : "";

  const resetLoad = () => {
    setTrips(null);
    setResult(null);
    clear();
  };
  const pick = (id: string) => {
    setEmpId(id);
    setTaxCode(db.employees.find((e) => e.id === id)?.taxCode ?? "");
    resetLoad();
  };
  const doLoad = async () => {
    setBusy("load");
    setField(undefined);
    setResult(null);
    try {
      const t = await loadTrips(empId, year, month, dbName);
      setTrips(t);
      setStatus(
        t.length
          ? {
              tone: "info",
              title: `Loaded ${t.length} trips for ${who}, ${MONTHS[month - 1]} ${year}.`,
            }
          : null,
      );
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Trips not loaded", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const doSave = async () => {
    setBusy("save");
    setField(undefined);
    try {
      const p = await saveCommissionPayment({
        empId,
        year,
        month,
        commission,
        suretyFund: fund,
        ceiling: ceiling ?? 0,
        beginningBalance: beginning ?? 0,
        taxAmount: tax ?? 0,
        taxCode,
      });
      setResult(p);
      setStatus(
        p.contributionsUpdated
          ? {
              tone: "success",
              title: `Saved payment for ${who}, ${MONTHS[month - 1]} ${year}.`,
              body: "Any earlier payment for this month was replaced. SSS, PhilHealth, Pag-IBIG, and tax were written onto it.",
            }
          : {
              tone: "error",
              title: `Saved payment for ${who}, ${MONTHS[month - 1]} ${year}, but contributions were not updated.`,
              body: "Commission is 0, so the server did not write SSS, PhilHealth, Pag-IBIG, or tax.",
            },
      );
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Payment not saved", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const err = (f: string) => (field === f ? (status?.body as string) : null);

  const payments = db.payments
    .filter(
      (p) =>
        p.empId === empId &&
        p.year === listYear &&
        (!quarter || Math.ceil(p.month / 3) === Number(quarter)),
    )
    .sort((a, b) => b.month - a.month);

  return (
    <>
      <PageHeader
        section="Payroll"
        title="Driver commissions"
        description="FMS only. Driver money comes from trips in the FMS database, not from hours. Separate from payroll entry."
      />

      <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1fr)_380px]">
        <div className="min-w-0 space-y-4">
          <Panel>
            <PanelHeader
              title="Trips"
              actions={
                <Button
                  variant={trips ? "secondary" : "primary"}
                  onClick={doLoad}
                  busy={busy === "load"}
                  disabled={!!busy || !empId}
                >
                  Load trips
                </Button>
              }
            />
            <div className="grid gap-3 border-b border-grey-100 px-4 py-3 sm:grid-cols-3 lg:grid-cols-6">
              <Field label="Driver or conductor">
                <Select
                  value={role}
                  onChange={(e) => {
                    const r = e.target.value as "Driver" | "Conductor";
                    setRole(r);
                    pick(
                      db.employees.find(
                        (x) => fmsIds.has(x.clientId) && !isExcluded(x) && x.role === r,
                      )?.id ?? "",
                    );
                  }}
                >
                  <option>Driver</option>
                  <option>Conductor</option>
                </Select>
              </Field>
              <Field label="Employee" className="lg:col-span-2" error={err("empId")}>
                <Select value={empId} onChange={(e) => pick(e.target.value)}>
                  {people.length === 0 && <option value="">None</option>}
                  {people.map((e) => (
                    <option key={e.id} value={e.id}>
                      {e.empNo} · {fullName(e)}
                    </option>
                  ))}
                </Select>
              </Field>
              <Field label="Department">
                <Readout>{emp?.department ?? "—"}</Readout>
              </Field>
              <Field label="Year">
                <Select
                  value={year}
                  onChange={(e) => (setYear(Number(e.target.value)), resetLoad())}
                >
                  {[2025, 2024].map((y) => (
                    <option key={y}>{y}</option>
                  ))}
                </Select>
              </Field>
              <Field label="Month">
                <Select
                  value={month}
                  onChange={(e) => (setMonth(Number(e.target.value)), resetLoad())}
                >
                  {MONTHS.map((m, i) => (
                    <option key={m} value={i + 1}>
                      {m}
                    </option>
                  ))}
                </Select>
              </Field>
              <Field label="FMS database name" className="lg:col-span-2" error={err("dbName")}>
                <TextInput
                  value={dbName}
                  onChange={(e) => (setDbName(e.target.value), resetLoad())}
                />
              </Field>
              <Field label="Tax code" error={err("taxCode")}>
                <Select value={taxCode} onChange={(e) => setTaxCode(e.target.value)}>
                  <option value="">Choose</option>
                  {db.taxCodes.map((t) => (
                    <option key={t.id} value={t.code}>
                      {t.code}
                    </option>
                  ))}
                </Select>
              </Field>
            </div>
            {status && (
              <div className="border-b border-grey-100 px-4 py-3">
                <StatusLine status={status} onDismiss={clear} />
              </div>
            )}
            {trips === null ? (
              <EmptyState>Choose the person, month, and FMS database, then Load trips.</EmptyState>
            ) : (
              <DataTable
                rows={trips}
                rowKey={(t) => t.id}
                dense
                maxHeight="420px"
                empty={
                  <EmptyState>{`${dbName} has no trips for ${who} in ${MONTHS[month - 1]} ${year}.`}</EmptyState>
                }
                columns={[
                  { key: "d", header: "Date", render: (t) => fmtDate(t.date) },
                  { key: "b", header: "Bundle", render: (t) => t.bundle },
                  {
                    key: "c",
                    header: "Commission",
                    align: "right",
                    render: (t) => <Amount value={t.commission} />,
                  },
                  {
                    key: "f",
                    header: "Fund",
                    align: "right",
                    render: (t) => <Amount value={t.fund} />,
                  },
                  { key: "bus", header: "Bus", render: (t) => t.bus },
                  {
                    key: "r",
                    header: "Route",
                    render: (t) => <span className="text-grey-600">{t.route}</span>,
                  },
                ]}
              />
            )}
          </Panel>

          <Panel>
            <PanelHeader
              title="Saved payments"
              meta={who}
              actions={
                <>
                  <Select
                    aria-label="Quarter"
                    value={quarter}
                    onChange={(e) => setQuarter(e.target.value)}
                    className="w-36"
                  >
                    <option value="">All quarters</option>
                    {[1, 2, 3, 4].map((q) => (
                      <option key={q} value={q}>
                        Quarter {q}
                      </option>
                    ))}
                  </Select>
                  <Select
                    aria-label="Year"
                    value={listYear}
                    onChange={(e) => setListYear(Number(e.target.value))}
                    className="w-24"
                  >
                    {[2025, 2024].map((y) => (
                      <option key={y}>{y}</option>
                    ))}
                  </Select>
                </>
              }
            />
            <DataTable
              rows={payments}
              rowKey={(p) => p.id}
              empty={
                <EmptyState>
                  No saved payments for this person in that period. Load trips, then Save payment.
                </EmptyState>
              }
              columns={[
                { key: "m", header: "Month", render: (p) => `${MONTHS[p.month - 1]} ${p.year}` },
                {
                  key: "c",
                  header: "Commission",
                  align: "right",
                  render: (p) => <Amount value={p.commission} />,
                },
                {
                  key: "s",
                  header: "SSS",
                  align: "right",
                  render: (p) => <Amount value={p.sssEe} />,
                },
                {
                  key: "ph",
                  header: "PhilHealth",
                  align: "right",
                  render: (p) => <Amount value={p.phEe} />,
                },
                {
                  key: "hd",
                  header: "Pag-IBIG",
                  align: "right",
                  render: (p) => <Amount value={p.hdmfEe} />,
                },
                {
                  key: "t",
                  header: "Tax amount",
                  align: "right",
                  render: (p) => <Amount value={p.taxAmount} />,
                },
                {
                  key: "u",
                  header: "Contributions",
                  render: (p) =>
                    p.contributionsUpdated ? (
                      <span className="text-cyan-800">Updated</span>
                    ) : (
                      <span className="font-medium text-red-600">Not updated</span>
                    ),
                },
                {
                  key: "by",
                  header: "Saved",
                  render: (p) => (
                    <span className="text-grey-600">{`${p.savedBy}, ${fmtDateTime(p.savedAt)}`}</span>
                  ),
                },
              ]}
            />
          </Panel>
        </div>

        <Panel className="xl:sticky xl:top-[7.5rem]">
          <PanelHeader
            title="Payment"
            meta={`${MONTHS[month - 1]} ${year}`}
            actions={
              <Button
                variant={trips ? "primary" : "secondary"}
                onClick={doSave}
                busy={busy === "save"}
                disabled={!trips || !!busy}
              >
                Save payment
              </Button>
            }
          />
          <div className="space-y-3 px-4 py-4">
            <table className="w-full text-sm">
              <tbody>
                <SumRow label="Commission" value={money(commission)} strong />
                <SumRow label="Surety fund" value={money(fund)} />
              </tbody>
            </table>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Ceiling" error={err("ceiling")}>
                <NumberInput value={ceiling} onValue={setCeiling} />
              </Field>
              <Field label="Beginning balance">
                <NumberInput value={beginning} onValue={setBeginning} />
              </Field>
              <Field
                label="Tax amount"
                className="col-span-2"
                hint="Starts at 200.00. Change it if needed; it is not computed."
              >
                <NumberInput value={tax} onValue={setTax} />
              </Field>
            </div>
            {capped && (
              <Notice
                tone="info"
                title={`Contributions use ${money(ceiling)}, not ${money(commission)}.`}
              >
                Commission is above the ceiling.
              </Notice>
            )}
            <p className="rounded-ctl bg-grey-50 px-3 py-2 text-xs leading-5 text-grey-600">
              Commission above the ceiling is cut to the ceiling before contributions. Tax is the
              amount in the tax box. Saving replaces any payment already stored for this person,
              month, and year.
            </p>
          </div>
          {result && (
            <div className="border-t border-grey-100">
              <div className="bg-grey-50 px-4 py-2 text-sm font-semibold text-grey-900">Result</div>
              {result.contributionsUpdated ? (
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-xs text-grey-500">
                      <th className="h-8 px-4 text-left font-medium" />
                      <th className="px-3 text-right font-medium">Employee</th>
                      <th className="px-4 text-right font-medium">Employer</th>
                    </tr>
                  </thead>
                  <tbody>
                    <ShareRow label="SSS" ee={result.sssEe} er={result.sssEr} />
                    <ShareRow label="PhilHealth" ee={result.phEe} er={result.phEr} />
                    <ShareRow label="Pag-IBIG" ee={result.hdmfEe} er={result.hdmfEr} />
                    <ShareRow label="ECC" ee={null} er={result.ecc} />
                    <ShareRow label="Tax amount" ee={result.taxAmount} er={null} />
                  </tbody>
                </table>
              ) : (
                <p className="px-4 py-3 text-sm font-medium text-red-600">
                  Contributions were not updated: commission is 0.
                </p>
              )}
            </div>
          )}
        </Panel>
      </div>
    </>
  );
}

function SumRow({ label, value, strong }: { label: string; value: string; strong?: boolean }) {
  return (
    <tr className="border-b border-grey-100">
      <td className="h-9 text-grey-700">{label}</td>
      <td
        className={
          strong ? "text-right text-base font-semibold text-grey-900" : "text-right text-grey-900"
        }
      >
        {value}
      </td>
    </tr>
  );
}
function ShareRow({ label, ee, er }: { label: string; ee: number | null; er: number | null }) {
  return (
    <tr className="border-t border-grey-100">
      <td className="h-9 px-4 text-grey-700">{label}</td>
      <td className="px-3 text-right">
        {ee == null ? <span className="text-grey-300">—</span> : <Amount value={ee} />}
      </td>
      <td className="px-4 text-right">
        {er == null ? <span className="text-grey-300">—</span> : <Amount value={er} />}
      </td>
    </tr>
  );
}
