import { useState } from "react";
import { useSession } from "../app/session";
import { ApiError, deleteDeduction, saveDeduction } from "../data/api";
import type { Deduction } from "../data/types";
import { DEDUCTION_CODES, deductionLabel, fullName, isExcluded, isFmsClient } from "../data/rules";
import { uid } from "../data/money";
import {
  Amount,
  Button,
  Checkbox,
  DataTable,
  EmptyState,
  Field,
  Notice,
  NumberInput,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  StatusLine,
  Tag,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { Icon } from "../ui/icons";

export function Deductions() {
  const { db, clients, cutoff, can } = useSession();
  const confirm = useConfirm();
  const ids = new Set(clients.map((c) => c.id));
  const people = db.employees
    .filter((e) => ids.has(e.clientId) && !isExcluded(e))
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const [empId, setEmpId] = useState(
    people.find((e) => e.clientId === cutoff?.clientId)?.id ?? people[0]?.id ?? "",
  );
  const [draft, setDraft] = useState<Deduction | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [field, setField] = useState<string | undefined>();
  const [busy, setBusy] = useState<string | null>(null);
  const { status, setStatus, clear } = useStatus();

  const emp = db.employees.find((e) => e.id === empId);
  const fms = isFmsClient(db.clients.find((c) => c.id === emp?.clientId));
  const rows = db.deductions.filter((d) => d.empId === empId).sort((a, b) => a.code - b.code);
  const saved = draft ? db.deductions.find((d) => d.id === draft.id) : undefined;
  const locked = !!saved?.locked;
  const editable = (isNew ? can.add : can.edit) && !locked;
  const who = emp ? `${emp.empNo} ${fullName(emp)}` : "";

  const open = (d: Deduction) => {
    setDraft({ ...d });
    setIsNew(false);
    setField(undefined);
    clear();
  };
  const startNew = () => {
    setDraft({
      id: uid(),
      empId,
      code: 0,
      balance: 0,
      schedule: "Every cutoff",
      amount: 0,
      locked: false,
    });
    setIsNew(true);
    setField(undefined);
    clear();
  };
  const save = async () => {
    if (!draft) return;
    setBusy("save");
    setField(undefined);
    try {
      await saveDeduction(draft);
      setIsNew(false);
      setStatus({
        tone: "success",
        title: `Saved ${deductionLabel(draft.code, fms).toLowerCase()} for ${who}.`,
      });
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Deduction not saved", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const remove = async () => {
    if (!saved) return;
    const ok = await confirm({
      title: "Delete deduction",
      body: (
        <>
          Delete <b>{deductionLabel(saved.code, fms)}</b> for {who}? Payroll stops taking it.
        </>
      ),
      confirmLabel: "Delete",
      tone: "danger",
    });
    if (!ok) return;
    setBusy("delete");
    try {
      await deleteDeduction(saved.id);
      setDraft(null);
      setStatus({
        tone: "success",
        title: `Deleted ${deductionLabel(saved.code, fms).toLowerCase()} for ${who}.`,
      });
    } catch (e) {
      setStatus({ tone: "error", title: "Deduction not deleted", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const set = <K extends keyof Deduction>(k: K, v: Deduction[K]) =>
    setDraft((d) => (d ? { ...d, [k]: v } : d));
  const err = (f: string) => (field === f ? (status?.body as string) : null);

  return (
    <>
      <PageHeader
        section="Time keeping"
        title="Deductions"
        description="Loans and recurring deductions payroll takes each cutoff."
      />
      <Panel className="mb-4">
        <div className="flex flex-wrap items-end gap-4 px-4 py-3">
          <Field label="Employee" className="w-[360px] max-w-full">
            <Select
              value={empId}
              onChange={(e) => {
                setEmpId(e.target.value);
                setDraft(null);
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
          <div className="flex items-center gap-2 pb-1.5 text-sm text-grey-600">
            {db.clients.find((c) => c.id === emp?.clientId)?.name}
            {fms && <Tag tone="blue">FMS labels</Tag>}
          </div>
        </div>
      </Panel>

      <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1.3fr)_minmax(360px,1fr)]">
        <Panel>
          <PanelHeader
            title="Deductions"
            meta={`${rows.length} for this person`}
            actions={
              can.add && (
                <Button onClick={startNew} disabled={!!busy || !empId}>
                  New deduction
                </Button>
              )
            }
          />
          <DataTable
            rows={rows}
            rowKey={(d) => d.id}
            selected={draft?.id}
            onSelect={open}
            empty={
              <EmptyState
                action={can.add ? <Button onClick={startNew}>New deduction</Button> : undefined}
              >
                No deductions for this person.
              </EmptyState>
            }
            columns={[
              {
                key: "c",
                header: "Code",
                width: "56px",
                render: (d) => <span className="font-medium">{d.code}</span>,
              },
              { key: "n", header: "Name", render: (d) => deductionLabel(d.code, fms) },
              {
                key: "b",
                header: "Balance",
                align: "right",
                render: (d) => <Amount value={d.balance} />,
              },
              {
                key: "s",
                header: "Schedule",
                render: (d) => <span className="text-grey-600">{d.schedule}</span>,
              },
              {
                key: "a",
                header: "This cutoff",
                align: "right",
                render: (d) => <Amount value={d.amount} strong />,
              },
              {
                key: "l",
                header: "Locked",
                render: (d) =>
                  d.locked ? (
                    <span className="inline-flex items-center gap-1 text-grey-700">
                      <Icon.lock className="size-3.5" />
                      Locked
                    </span>
                  ) : (
                    <span className="text-grey-400">No</span>
                  ),
              },
            ]}
          />
        </Panel>

        <Panel>
          {!draft ? (
            <>
              <PanelHeader title="No deduction open" />
              <div className="space-y-3 px-4 py-4">
                <StatusLine status={status} onDismiss={clear} />
                <p className="text-sm text-grey-600">
                  Choose a deduction to edit it{can.add ? ", or use New deduction" : ""}.
                </p>
              </div>
            </>
          ) : (
            <>
              <PanelHeader
                title={isNew ? "New deduction" : deductionLabel(draft.code, fms)}
                meta={locked ? "Locked" : undefined}
                actions={
                  <>
                    {!isNew && can.del && (
                      <Button
                        variant="danger"
                        onClick={remove}
                        disabled={locked || !!busy}
                        busy={busy === "delete"}
                      >
                        Delete
                      </Button>
                    )}
                    {(isNew ? can.add : can.edit) && (
                      <>
                        <Button
                          onClick={() => (isNew ? setDraft(null) : saved && open(saved))}
                          disabled={locked || !!busy}
                        >
                          Cancel
                        </Button>
                        <Button
                          variant="primary"
                          onClick={save}
                          disabled={locked || !!busy}
                          busy={busy === "save"}
                        >
                          Save
                        </Button>
                      </>
                    )}
                  </>
                }
              />
              <div className="space-y-4 px-4 py-4">
                <StatusLine status={status} onDismiss={clear} />
                {locked && (
                  <Notice tone="neutral" title="This deduction is locked">
                    Save and Delete are disabled while it is locked.
                  </Notice>
                )}
                <div className="grid grid-cols-2 gap-3">
                  <Field label="Code" required error={err("code")} className="col-span-2">
                    <Select
                      value={String(draft.code || "")}
                      disabled={!editable}
                      onChange={(e) => set("code", Number(e.target.value))}
                    >
                      <option value="">Choose a code</option>
                      {DEDUCTION_CODES.map((c) => (
                        <option key={c.code} value={c.code}>
                          {c.code} · {fms && c.fms ? c.fms : c.staff}
                        </option>
                      ))}
                    </Select>
                  </Field>
                  <Field label="Balance">
                    <NumberInput
                      value={draft.balance}
                      disabled={!editable}
                      onValue={(n) => set("balance", n ?? 0)}
                    />
                  </Field>
                  <Field label="Amount this cutoff" error={err("amount")}>
                    <NumberInput
                      value={draft.amount}
                      disabled={!editable}
                      onValue={(n) => set("amount", n ?? 0)}
                    />
                  </Field>
                  <Field label="Schedule" className="col-span-2">
                    <Select
                      value={draft.schedule}
                      disabled={!editable}
                      onChange={(e) => set("schedule", e.target.value as Deduction["schedule"])}
                    >
                      <option>Every cutoff</option>
                      <option>First half</option>
                      <option>Second half</option>
                    </Select>
                  </Field>
                </div>
                <Checkbox
                  checked={draft.locked}
                  disabled={!editable}
                  onChange={(v) => set("locked", v)}
                  label="Locked"
                  hint="A locked deduction cannot be changed or deleted."
                />
              </div>
            </>
          )}
        </Panel>
      </div>
    </>
  );
}
