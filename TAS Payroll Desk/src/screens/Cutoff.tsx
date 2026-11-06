import { useState } from "react";
import { useSession } from "../app/session";
import { ApiError, cutoffRecord, selectCutoff } from "../data/api";
import {
  cutoffHalf,
  cutoffKey,
  isFmsClient,
  showsSssPhilhealth,
  showsUnionDues,
} from "../data/rules";
import {
  Button,
  DataTable,
  EmptyState,
  Field,
  KeyValues,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  StatusLine,
  Tag,
  TextInput,
  useStatus,
} from "../ui/kit";
import { Icon } from "../ui/icons";
import { fmtDate, fmtDateTime, fmtRange } from "../ui/format";
import { navigate } from "../app/router";
import { canSee } from "../data/rules";

export function CutoffScreen() {
  const { db, user, clients, cutoff } = useSession();
  const active = clients.filter((c) => c.active);
  const [clientId, setClientId] = useState(cutoff?.clientId ?? active[0]?.id ?? "");
  const [start, setStart] = useState(cutoff?.start ?? "");
  const [end, setEnd] = useState(cutoff?.end ?? "");
  const [busy, setBusy] = useState(false);
  const [field, setField] = useState<string | undefined>();
  const { status, setStatus, clear } = useStatus();

  const client = db.clients.find((c) => c.id === clientId);
  const validEnd = /^\d{4}-\d{2}-\d{2}$/.test(end);
  const rec = validEnd ? cutoffRecord(db, { clientId, start, end }) : undefined;
  const history = db.cutoffs
    .filter((c) => c.clientId === clientId)
    .sort((a, b) => b.end.localeCompare(a.end));
  const isSelected =
    cutoff && cutoff.clientId === clientId && cutoff.start === start && cutoff.end === end;

  const submit = async () => {
    setBusy(true);
    setField(undefined);
    try {
      await selectCutoff({ clientId, start, end });
      setStatus({
        tone: "success",
        title: `Selected cutoff ${client?.name}, ${fmtRange(start, end)}.`,
        body: "Time keeping and payroll entry now read this period.",
      });
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Cutoff not selected", body: (e as Error).message });
    } finally {
      setBusy(false);
    }
  };
  const err = (f: string) => (field === f ? (status?.body as string) : null);

  return (
    <>
      <PageHeader
        section="Payroll"
        title="Cutoff"
        description="Pick the period before time keeping or payroll. Both screens read this selection."
      />
      <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1fr)_minmax(0,1fr)]">
        <Panel>
          <PanelHeader
            title="Period"
            actions={
              <Button variant="primary" onClick={submit} busy={busy} disabled={!!isSelected}>
                {isSelected ? "Selected" : "Select cutoff"}
              </Button>
            }
          />
          <div className="space-y-4 px-4 py-4">
            <StatusLine status={status} onDismiss={clear} />
            <div className="grid gap-3 sm:grid-cols-3">
              <Field label="Client" required error={err("clientId")} className="sm:col-span-3">
                <Select value={clientId} onChange={(e) => setClientId(e.target.value)}>
                  {active.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name} {isFmsClient(c) ? "(FMS)" : "(Staff)"}
                    </option>
                  ))}
                </Select>
              </Field>
              <Field label="Start date" required error={err("start")}>
                <TextInput type="date" value={start} onChange={(e) => setStart(e.target.value)} />
              </Field>
              <Field
                label="End date"
                required
                error={err("end")}
                hint="The end day decides the half."
              >
                <TextInput type="date" value={end} onChange={(e) => setEnd(e.target.value)} />
              </Field>
            </div>

            <div className="rounded-ctl border border-grey-100 bg-grey-50 px-3 py-3">
              {validEnd ? (
                <KeyValues
                  cols={3}
                  items={[
                    { label: "Half", value: `Half ${cutoffHalf(end)}` },
                    { label: "DTR key", value: cutoffKey(end) },
                    {
                      label: "Billed",
                      value: rec?.billed ? (
                        <span className="inline-flex items-center gap-1 text-cyan-800">
                          <Icon.lock className="size-3.5" /> Yes, read-only
                        </span>
                      ) : (
                        "No"
                      ),
                    },
                    {
                      label: "Contributions this half",
                      value: showsSssPhilhealth(end) ? "SSS and PhilHealth" : "Pag-IBIG",
                    },
                    {
                      label: "Union dues",
                      value: showsUnionDues(end) ? "Deducted" : "Zero (end day after 15th)",
                    },
                    { label: "Tax", value: "Both halves" },
                  ]}
                />
              ) : (
                <p className="text-sm text-grey-600">
                  Enter an end date to see the half, DTR key, and which contributions apply.
                </p>
              )}
            </div>
            <p className="text-xs text-grey-500">
              A cutoff is marked billed only by Move to master on Payroll entry
              {rec?.billed && rec.billedBy
                ? `. This one was billed by ${rec.billedBy} on ${fmtDateTime(rec.billedAt)}.`
                : "."}
            </p>
            {isSelected && (
              <div className="flex flex-wrap gap-2 border-t border-grey-100 pt-4">
                {canSee(user, "timeKeeping") && (
                  <Button onClick={() => navigate(isFmsClient(client) ? "time-fms" : "time-staff")}>
                    Go to time keeping
                  </Button>
                )}
                {canSee(user, "payroll") && (
                  <Button onClick={() => navigate("payroll-entry")}>Go to payroll entry</Button>
                )}
              </div>
            )}
          </div>
        </Panel>

        <Panel>
          <PanelHeader title="Cutoffs for this client" meta="Choose one to fill the period" />
          <DataTable
            rows={history}
            rowKey={(r) => r.key}
            selected={validEnd ? cutoffKey(end) : null}
            onSelect={(r) => {
              setStart(r.start);
              setEnd(r.end);
              clear();
            }}
            empty={
              <EmptyState>
                No cutoffs yet for this client. Enter a start and end date to create the first one.
              </EmptyState>
            }
            columns={[
              {
                key: "key",
                header: "DTR key",
                render: (r) => <span className="font-medium">{r.key}</span>,
              },
              { key: "s", header: "Start", render: (r) => fmtDate(r.start) },
              { key: "e", header: "End", render: (r) => fmtDate(r.end) },
              { key: "h", header: "Half", align: "right", render: (r) => cutoffHalf(r.end) },
              {
                key: "b",
                header: "Status",
                render: (r) =>
                  r.billed ? <Tag tone="success">Billed</Tag> : <Tag tone="outline">Open</Tag>,
              },
            ]}
          />
        </Panel>
      </div>
    </>
  );
}
