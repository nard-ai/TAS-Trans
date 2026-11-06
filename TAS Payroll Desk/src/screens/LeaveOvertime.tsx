import { useState } from "react";
import { useSession } from "../app/session";
import { ApiError, deleteLeave, saveLeave, saveOvertime } from "../data/api";
import { LEAVE_LETTERS, fullName, leaveLabel } from "../data/rules";
import {
  Button,
  DataTable,
  EmptyState,
  Field,
  NumberInput,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  StatusLine,
  TextInput,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { fmtDate, hrs } from "../ui/format";
import { Icon } from "../ui/icons";

export function LeaveOvertime() {
  const { db, clients, cutoff, can } = useSession();
  const confirm = useConfirm();
  const ids = new Set(clients.map((c) => c.id));
  const people = db.employees
    .filter((e) => ids.has(e.clientId) && e.active)
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const firstInCutoff = people.find((e) => e.clientId === cutoff?.clientId);
  const [empId, setEmpId] = useState(firstInCutoff?.id ?? people[0]?.id ?? "");
  const [selLeave, setSelLeave] = useState<string | null>(null);
  const [date, setDate] = useState(cutoff?.start ?? "");
  const [letter, setLetter] = useState("");
  const [hours, setHours] = useState<number | null>(8);
  const [otDate, setOtDate] = useState(cutoff?.start ?? "");
  const [otHours, setOtHours] = useState<number | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [field, setField] = useState<string | undefined>();
  const { status, setStatus, clear } = useStatus();

  const emp = db.employees.find((e) => e.id === empId);
  const leaves = db.leaves
    .filter((l) => l.empId === empId)
    .sort((a, b) => b.date.localeCompare(a.date));
  const ots = db.overtimes
    .filter((o) => o.empId === empId)
    .sort((a, b) => b.date.localeCompare(a.date));
  const balances = db.leaveBalances.filter((b) => b.empId === empId);

  const run = async (key: string, fn: () => Promise<void>, ok: string) => {
    setBusy(key);
    setField(undefined);
    try {
      await fn();
      setStatus({
        tone: "success",
        title: ok,
        body: "Use Refresh from leave and overtime on time keeping to apply it to computed rows.",
      });
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Not saved", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const err = (f: string) => (field === f ? (status?.body as string) : null);
  const who = emp ? `${emp.empNo} ${fullName(emp)}` : "";

  const onDelete = async () => {
    const l = leaves.find((x) => x.id === selLeave);
    if (!l) return;
    const ok = await confirm({
      title: "Delete leave",
      body: (
        <>
          Delete {leaveLabel(l.letter).toLowerCase()} on <b>{fmtDate(l.date)}</b> for {who}?
        </>
      ),
      confirmLabel: "Delete",
      tone: "danger",
    });
    if (!ok) return;
    await run("delete", () => deleteLeave(l.id), `Deleted leave on ${fmtDate(l.date)} for ${who}.`);
    setSelLeave(null);
  };

  const printBalance = () => window.print();

  return (
    <>
      <PageHeader
        section="Time keeping"
        title="Leave and overtime"
        description="Per employee, per date: a leave letter and hours. Filed overtime replaces the computed overtime for that day."
        actions={
          <Button onClick={printBalance} icon={<Icon.printer />}>
            Print leave balance
          </Button>
        }
      />
      <Panel className="no-print mb-4">
        <div className="flex flex-wrap items-end gap-4 px-4 py-3">
          <Field label="Employee" className="w-[360px] max-w-full">
            <Select
              value={empId}
              onChange={(e) => {
                setEmpId(e.target.value);
                setSelLeave(null);
                clear();
              }}
            >
              {people.map((e) => (
                <option key={e.id} value={e.id}>
                  {e.empNo} · {fullName(e)}
                </option>
              ))}
            </Select>
          </Field>
          <div className="pb-1.5 text-sm text-grey-600">
            {balances.length
              ? balances
                  .map((b) => `${b.type}: ${(b.entitled - b.used).toFixed(1)} days left`)
                  .join(" · ")
              : "No leave balances recorded."}
          </div>
        </div>
      </Panel>

      <StatusLine status={status} onDismiss={clear} className="no-print mb-4" />

      <div className="no-print grid items-start gap-4 xl:grid-cols-[minmax(0,1.3fr)_minmax(0,1fr)]">
        <Panel>
          <PanelHeader
            title="Leave and day off"
            meta={`${leaves.length} filed`}
            actions={
              <>
                {can.del && (
                  <Button
                    variant="danger"
                    onClick={onDelete}
                    disabled={!selLeave || !!busy}
                    busy={busy === "delete"}
                  >
                    Delete leave
                  </Button>
                )}
                {(can.add || can.edit) && (
                  <Button
                    variant="primary"
                    busy={busy === "leave"}
                    disabled={!!busy}
                    onClick={() =>
                      run(
                        "leave",
                        () => saveLeave({ empId, date, letter, hours: hours ?? 0 }),
                        `Saved ${leaveLabel(letter).toLowerCase()} on ${fmtDate(date)} for ${who}.`,
                      )
                    }
                  >
                    Save leave
                  </Button>
                )}
              </>
            }
          />
          <div className="grid gap-3 border-b border-grey-100 px-4 py-3 sm:grid-cols-[1fr_1.4fr_0.8fr]">
            <Field label="Date" required error={err("date")}>
              <TextInput type="date" value={date} onChange={(e) => setDate(e.target.value)} />
            </Field>
            <Field label="Letter" required error={err("letter")}>
              <Select value={letter} onChange={(e) => setLetter(e.target.value)}>
                <option value="">Choose a letter</option>
                {LEAVE_LETTERS.map((l) => (
                  <option key={l.letter} value={l.letter}>
                    {l.letter} · {l.label}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Hours" required error={err("hours")}>
              <NumberInput value={hours} onValue={setHours} />
            </Field>
            {letter && (
              <p className="text-xs text-grey-500 sm:col-span-3">
                Payroll reads {letter} as:{" "}
                {LEAVE_LETTERS.find((l) => l.letter === letter)?.reads.toLowerCase()}.
              </p>
            )}
          </div>
          <DataTable
            rows={leaves}
            rowKey={(l) => l.id}
            selected={selLeave}
            onSelect={(l) => {
              setSelLeave(l.id);
              setDate(l.date);
              setLetter(l.letter);
              setHours(l.hours);
            }}
            maxHeight="360px"
            empty={
              <EmptyState>
                No leave filed for this person. Fill the date, letter, and hours above, then Save
                leave.
              </EmptyState>
            }
            columns={[
              { key: "d", header: "Date", render: (l) => fmtDate(l.date) },
              {
                key: "l",
                header: "Letter",
                width: "60px",
                render: (l) => <span className="font-semibold">{l.letter}</span>,
              },
              { key: "m", header: "Screen label", render: (l) => leaveLabel(l.letter) },
              { key: "h", header: "Hours", align: "right", render: (l) => hrs(l.hours) },
            ]}
          />
        </Panel>

        <Panel>
          <PanelHeader
            title="Filed overtime"
            meta={`${ots.length} filed`}
            actions={
              (can.add || can.edit) && (
                <Button
                  busy={busy === "ot"}
                  disabled={!!busy}
                  onClick={() =>
                    run(
                      "ot",
                      () => saveOvertime({ empId, date: otDate, hours: otHours ?? -1 }),
                      `Saved ${hrs(otHours)} overtime hours on ${fmtDate(otDate)} for ${who}.`,
                    )
                  }
                >
                  Save overtime
                </Button>
              )
            }
          />
          <div className="grid grid-cols-2 gap-3 border-b border-grey-100 px-4 py-3">
            <Field label="Date" required error={err("otDate")}>
              <TextInput type="date" value={otDate} onChange={(e) => setOtDate(e.target.value)} />
            </Field>
            <Field
              label="Overtime hours"
              required
              error={err("otHours")}
              hint="Replaces computed overtime for the day."
            >
              <NumberInput value={otHours} onValue={setOtHours} />
            </Field>
          </div>
          <DataTable
            rows={ots}
            rowKey={(o) => o.id}
            onSelect={(o) => {
              setOtDate(o.date);
              setOtHours(o.hours);
            }}
            empty={<EmptyState>No overtime filed for this person.</EmptyState>}
            columns={[
              { key: "d", header: "Date", render: (o) => fmtDate(o.date) },
              { key: "h", header: "Hours", align: "right", render: (o) => hrs(o.hours) },
            ]}
          />
        </Panel>
      </div>

      <div className="print-only">
        <h1 className="text-lg font-semibold">{db.company?.name ?? ""}</h1>
        <h2 className="mb-3 text-base font-semibold">Leave balance: {who}</h2>
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b">
              <th className="py-1 text-left">Year</th>
              <th className="text-left">Type</th>
              <th className="text-right">Entitled</th>
              <th className="text-right">Used</th>
              <th className="text-right">Balance</th>
            </tr>
          </thead>
          <tbody>
            {balances.map((b) => (
              <tr key={b.id} className="border-b">
                <td className="py-1">{b.year}</td>
                <td>{b.type}</td>
                <td className="text-right">{b.entitled.toFixed(1)}</td>
                <td className="text-right">{b.used.toFixed(1)}</td>
                <td className="text-right">{(b.entitled - b.used).toFixed(1)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}
