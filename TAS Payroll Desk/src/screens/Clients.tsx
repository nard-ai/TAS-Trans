import { useMemo, useState } from "react";
import { useSession } from "../app/session";
import { ApiError, saveClient } from "../data/api";
import type { Client, ClientRates } from "../data/types";
import { clientScope, isFmsClient } from "../data/rules";
import { uid } from "../data/money";
import {
  Button,
  Checkbox,
  DataTable,
  EmptyState,
  Field,
  FormSection,
  Notice,
  NumberInput,
  PageHeader,
  Panel,
  PanelHeader,
  Readout,
  SearchInput,
  Select,
  StatusLine,
  Tag,
  TextInput,
  useConfirm,
  useStatus,
} from "../ui/kit";

const RATE_FIELDS: { key: keyof ClientRates; label: string; unit: string }[] = [
  { key: "overtime", label: "Overtime", unit: "× hourly" },
  { key: "nightDiff", label: "Night differential", unit: "× hourly" },
  { key: "restDay", label: "Rest day", unit: "× hourly" },
  { key: "specialHoliday", label: "Special holiday", unit: "× hourly" },
  { key: "legalHoliday", label: "Legal holiday", unit: "× hourly" },
  { key: "holidayOvertime", label: "Holiday overtime", unit: "× hourly" },
  { key: "excessOvertime", label: "Excess overtime", unit: "× hourly" },
  { key: "ecola", label: "ECOLA", unit: "per day" },
  { key: "hazard", label: "Hazard", unit: "per day" },
  { key: "agencyRate", label: "Agency rate", unit: "percent" },
  { key: "dailyCharge", label: "Daily charge", unit: "per day" },
];

function blankClient(fms: boolean): Client {
  return {
    id: `c${uid()}`,
    code: "",
    code2: fms ? "FMS" : "",
    name: "",
    active: true,
    workingDays: null,
    rates: {
      overtime: 1.25,
      nightDiff: 0.1,
      restDay: 1.3,
      specialHoliday: 1.3,
      legalHoliday: 2,
      holidayOvertime: 2.6,
      excessOvertime: 1.69,
      ecola: 0,
      hazard: 0,
      agencyRate: 0,
      dailyCharge: 0,
    },
    flags: {
      sss: true,
      philhealth: true,
      pagibig: true,
      thirteenthOnCutoff: false,
      displayThirteenth: false,
      tax2018: true,
    },
    taxPercent: false,
    taxCeiling: 0,
  };
}

export function Clients() {
  const { db, user, can, clients } = useSession();
  const confirm = useConfirm();
  const scope = clientScope(user);
  const [q, setQ] = useState("");
  const [activeOnly, setActiveOnly] = useState(true);
  const [draft, setDraft] = useState<Client | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [field, setField] = useState<string | undefined>();
  const [busy, setBusy] = useState<string | null>(null);
  const { status, setStatus, clear } = useStatus();

  const list = useMemo(
    () =>
      clients
        .filter(
          (c) => (!activeOnly || c.active) && c.name.toLowerCase().includes(q.trim().toLowerCase()),
        )
        .sort((a, b) => a.code.localeCompare(b.code)),
    [clients, q, activeOnly],
  );
  const saved = draft ? db.clients.find((c) => c.id === draft.id) : undefined;
  const editable = isNew ? can.add : can.edit;

  const open = (c: Client) => {
    setDraft(structuredClone(c));
    setIsNew(false);
    setField(undefined);
    clear();
  };
  const startNew = () => {
    setDraft(blankClient(scope === "fms"));
    setIsNew(true);
    setField(undefined);
    clear();
  };
  const save = async (next = draft) => {
    if (!next) return;
    setBusy("save");
    setField(undefined);
    try {
      await saveClient(next);
      setIsNew(false);
      setDraft(structuredClone(next));
      setStatus({ tone: "success", title: `Saved client ${next.code} ${next.name}.` });
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Client not saved", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const deactivate = async () => {
    if (!draft) return;
    const ok = await confirm({
      title: "Deactivate client",
      body: (
        <>
          Set <b>{draft.name}</b> to inactive? It stops appearing for new cutoffs. Payroll history
          is kept.
        </>
      ),
      confirmLabel: "Deactivate",
    });
    if (!ok) return;
    const next = { ...draft, active: false };
    setDraft(next);
    setBusy("deactivate");
    await saveClient(next);
    setBusy(null);
    setStatus({
      tone: "success",
      title: `Deactivated client ${next.code} ${next.name}. History is kept.`,
    });
  };

  const set = <K extends keyof Client>(k: K, v: Client[K]) =>
    setDraft((d) => (d ? { ...d, [k]: v } : d));
  const setRate = (k: keyof ClientRates, v: number) =>
    setDraft((d) => (d ? { ...d, rates: { ...d.rates, [k]: v } } : d));
  const setFlag = (k: keyof Client["flags"], v: boolean) =>
    setDraft((d) => (d ? { ...d, flags: { ...d.flags, [k]: v } } : d));
  const err = (f: string) => (field === f ? (status?.body as string) : null);

  return (
    <>
      <PageHeader
        section="People"
        title="Clients and rates"
        description={
          scope === "all"
            ? "One row per client. Staff or FMS, pay rates, contribution flags, and the tax table switch live here."
            : `Your account type shows ${scope === "fms" ? "FMS clients only" : "staff clients only"}.`
        }
      />
      <div className="grid items-start gap-4 xl:grid-cols-[400px_minmax(0,1fr)]">
        <Panel>
          <PanelHeader
            title="Clients"
            meta={`${list.length} shown`}
            actions={
              can.add && (
                <Button onClick={startNew} disabled={!!busy}>
                  New client
                </Button>
              )
            }
          />
          <div className="flex items-center gap-3 border-b border-grey-100 px-4 py-2.5">
            <SearchInput
              label="Search clients by name"
              value={q}
              onChange={setQ}
              placeholder="Search by name"
              className="flex-1"
            />
            <Checkbox checked={activeOnly} onChange={setActiveOnly} label="Active only" />
          </div>
          <DataTable
            rows={list}
            rowKey={(c) => c.id}
            selected={draft?.id}
            onSelect={open}
            maxHeight="calc(100vh - 300px)"
            empty={
              <EmptyState
                action={can.add ? <Button onClick={startNew}>New client</Button> : undefined}
              >
                {q ? "No client matches that name." : "No clients yet."}
              </EmptyState>
            }
            columns={[
              {
                key: "code",
                header: "Code",
                width: "64px",
                render: (c) => <span className="font-medium">{c.code}</span>,
              },
              {
                key: "name",
                header: "Name",
                className: "whitespace-normal",
                render: (c) => c.name,
              },
              {
                key: "type",
                header: "Type",
                render: (c) => (isFmsClient(c) ? <Tag tone="blue">FMS</Tag> : <Tag>Staff</Tag>),
              },
              {
                key: "st",
                header: "Status",
                render: (c) =>
                  c.active ? (
                    <span className="text-grey-700">Active</span>
                  ) : (
                    <span className="text-grey-400">Inactive</span>
                  ),
              },
            ]}
          />
        </Panel>

        <Panel>
          {!draft ? (
            <>
              <PanelHeader title="No client open" />
              <div className="space-y-3 px-4 py-4">
                <StatusLine status={status} onDismiss={clear} />
                <p className="text-sm text-grey-600">
                  Choose a client to edit its rates and flags.
                </p>
              </div>
            </>
          ) : (
            <>
              <PanelHeader
                title={
                  isNew ? "New client" : `${saved?.code ?? draft.code} ${saved?.name ?? draft.name}`
                }
                meta={!isNew && !draft.active ? "Inactive" : undefined}
                actions={
                  editable && (
                    <>
                      {!isNew && draft.active && (
                        <Button onClick={deactivate} busy={busy === "deactivate"} disabled={!!busy}>
                          Deactivate
                        </Button>
                      )}
                      <Button
                        onClick={() =>
                          isNew ? (setDraft(null), setIsNew(false)) : saved && open(saved)
                        }
                        disabled={!!busy}
                      >
                        Cancel
                      </Button>
                      <Button
                        variant="primary"
                        onClick={() => save()}
                        busy={busy === "save"}
                        disabled={!!busy}
                      >
                        Save
                      </Button>
                    </>
                  )
                }
              />
              {status && (
                <div className="px-4 pt-4">
                  <StatusLine status={status} onDismiss={clear} />
                </div>
              )}
              {!editable && (
                <div className="px-4 pt-4">
                  <Notice tone="neutral" title="Read only">
                    Your account cannot edit clients.
                  </Notice>
                </div>
              )}
              <fieldset disabled={!editable} className="contents">
                <FormSection title="Client">
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field label="Code" required error={err("code")}>
                      <TextInput
                        value={draft.code}
                        onChange={(e) => set("code", e.target.value.toUpperCase())}
                      />
                    </Field>
                    <Field label="Name" required error={err("name")} className="sm:col-span-2">
                      <TextInput value={draft.name} onChange={(e) => set("name", e.target.value)} />
                    </Field>
                    <Field
                      label="Staff or FMS"
                      error={err("code2")}
                      hint="FMS means drivers and conductors."
                    >
                      <Select
                        value={draft.code2 === "FMS" ? "FMS" : ""}
                        onChange={(e) => set("code2", e.target.value)}
                        disabled={scope !== "all"}
                      >
                        <option value="">Staff</option>
                        <option value="FMS">FMS</option>
                      </Select>
                    </Field>
                  </div>
                  <div className="mt-3">
                    <Checkbox
                      checked={draft.active}
                      onChange={(v) => set("active", v)}
                      label="Active"
                      disabled={!editable}
                    />
                  </div>
                </FormSection>

                <FormSection title="Day basis">
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field
                      label="Working days per year"
                      hint="Blank means 303 on the payroll screen."
                    >
                      <NumberInput
                        value={draft.workingDays}
                        decimals={0}
                        placeholder="303"
                        onValue={(n) => set("workingDays", n)}
                      />
                    </Field>
                    <Field label="Daily hours basis" hint="Fixed for every client.">
                      <Readout>8 hours</Readout>
                    </Field>
                  </div>
                </FormSection>

                <FormSection
                  title="Rates"
                  description="Multipliers apply to the hourly rate. Amounts are in pesos."
                >
                  <div className="grid gap-x-3 gap-y-3 sm:grid-cols-3 lg:grid-cols-4">
                    {RATE_FIELDS.map((r) => (
                      <Field
                        key={r.key}
                        label={r.label}
                        aside={<span className="text-xs text-grey-400">{r.unit}</span>}
                      >
                        <NumberInput
                          value={draft.rates[r.key]}
                          onValue={(n) => setRate(r.key, n ?? 0)}
                        />
                      </Field>
                    ))}
                  </div>
                </FormSection>

                <FormSection title="Contributions and 13th month">
                  <div className="grid gap-2.5 sm:grid-cols-3">
                    <Checkbox
                      checked={draft.flags.sss}
                      disabled={!editable}
                      onChange={(v) => setFlag("sss", v)}
                      label="SSS"
                    />
                    <Checkbox
                      checked={draft.flags.philhealth}
                      disabled={!editable}
                      onChange={(v) => setFlag("philhealth", v)}
                      label="PhilHealth"
                    />
                    <Checkbox
                      checked={draft.flags.pagibig}
                      disabled={!editable}
                      onChange={(v) => setFlag("pagibig", v)}
                      label="Pag-IBIG"
                    />
                    <Checkbox
                      checked={draft.flags.thirteenthOnCutoff}
                      disabled={!editable}
                      onChange={(v) => setFlag("thirteenthOnCutoff", v)}
                      label="13th month on the cutoff"
                      hint="Adds daily rate / 12 on payroll entry."
                    />
                    <Checkbox
                      checked={draft.flags.displayThirteenth}
                      disabled={!editable}
                      onChange={(v) => setFlag("displayThirteenth", v)}
                      label="Display 13th month"
                    />
                  </div>
                </FormSection>

                <FormSection title="Tax">
                  <div className="grid items-start gap-3 sm:grid-cols-3">
                    <Checkbox
                      checked={draft.flags.tax2018}
                      disabled={!editable}
                      onChange={(v) => setFlag("tax2018", v)}
                      label="Use 2018 tax table"
                      hint="This same switch also pays ECOLA on rest days."
                    />
                    <Checkbox
                      checked={draft.taxPercent}
                      disabled={!editable}
                      onChange={(v) => set("taxPercent", v)}
                      label="Tax percentage"
                      hint="Tax is zero when the daily rate is at or under the ceiling."
                    />
                    <Field label="Daily-rate ceiling for zero tax">
                      <NumberInput
                        value={draft.taxCeiling}
                        disabled={!draft.taxPercent || !editable}
                        onValue={(n) => set("taxCeiling", n ?? 0)}
                      />
                    </Field>
                  </div>
                </FormSection>
              </fieldset>
            </>
          )}
        </Panel>
      </div>
    </>
  );
}
