import { useMemo, useState } from "react";
import { useSession } from "../app/session";
import { ApiError, copyEmployee, saveEmployee } from "../data/api";
import type { Employee } from "../data/types";
import { PAYROLL_STATUSES } from "../data/types";
import { EXCLUDED_STATUSES, fullName, isExcluded, isFmsClient } from "../data/rules";
import { uid } from "../data/money";
import {
  Button,
  Checkbox,
  DataTable,
  Dialog,
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
  cx,
  useStatus,
} from "../ui/kit";

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

export function Employees() {
  const { db, can, clients } = useSession();
  const clientIds = new Set(clients.map((c) => c.id));
  const [q, setQ] = useState("");
  const [showFilter, setShowFilter] = useState(false);
  const [clientFilter, setClientFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [draft, setDraft] = useState<Employee | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [field, setField] = useState<string | undefined>();
  const [busy, setBusy] = useState(false);
  const [copyOpen, setCopyOpen] = useState(false);
  const [copyNo, setCopyNo] = useState("");
  const [copyErr, setCopyErr] = useState<string | null>(null);
  const { status, setStatus, clear } = useStatus();

  const list = useMemo(() => {
    const term = q.trim().toLowerCase();
    return db.employees
      .filter((e) => clientIds.has(e.clientId))
      .filter((e) => !clientFilter || e.clientId === clientFilter)
      .filter(
        (e) => !statusFilter || (statusFilter === "excluded" ? isExcluded(e) : !isExcluded(e)),
      )
      .filter((e) => !term || e.empNo.includes(term) || fullName(e).toLowerCase().includes(term))
      .sort((a, b) => a.empNo.localeCompare(b.empNo));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [db.employees, q, clientFilter, statusFilter, clients]);

  const saved = draft ? db.employees.find((e) => e.id === draft.id) : undefined;
  const editable = isNew ? can.add : can.edit;
  const client = draft ? db.clients.find((c) => c.id === draft.clientId) : undefined;

  const open = (e: Employee) => {
    setDraft({ ...e });
    setIsNew(false);
    setField(undefined);
    clear();
  };
  const startNew = () => {
    setDraft({
      ...(saved ?? db.employees[0]),
      id: `e${uid()}`,
      empNo: "",
      dtrId: "",
      lastName: "",
      firstName: "",
      middleName: "",
      position: "",
      tin: "",
      sssNo: "",
      philhealthNo: "",
      pagibigNo: "",
      accountNo: "",
      phone: "",
      clientId: clientFilter || clients[0]?.id || "",
      payrollStatus: "A",
      active: true,
      hired: "",
      contractStart: "",
      contractEnd: "",
    });
    setIsNew(true);
    setField(undefined);
    clear();
  };
  const save = async () => {
    if (!draft) return;
    setBusy(true);
    setField(undefined);
    try {
      await saveEmployee(draft);
      setIsNew(false);
      setStatus({ tone: "success", title: `Saved employee ${draft.empNo} ${fullName(draft)}.` });
    } catch (e) {
      setField(e instanceof ApiError ? e.field : undefined);
      setStatus({ tone: "error", title: "Employee not saved", body: (e as Error).message });
    } finally {
      setBusy(false);
    }
  };
  const doCopy = async () => {
    if (!draft) return;
    setCopyErr(null);
    try {
      const copy = await copyEmployee(draft.id, copyNo);
      setCopyOpen(false);
      setCopyNo("");
      open(copy);
      setStatus({
        tone: "success",
        title: `Copied the profile onto employee ${copy.empNo}. Update the name before saving again.`,
      });
    } catch (e) {
      setCopyErr((e as Error).message);
    }
  };

  const set = <K extends keyof Employee>(k: K, v: Employee[K]) =>
    setDraft((d) => (d ? { ...d, [k]: v } : d));
  const err = (f: string) => (field === f ? (status?.body as string) : null);

  return (
    <>
      <PageHeader
        section="People"
        title="Employees"
        description="The 201 file and the pay profile payroll reads. Time records join on the DTR id."
      />
      <div className="grid items-start gap-4 xl:grid-cols-[440px_minmax(0,1fr)]">
        <Panel>
          <PanelHeader
            title="Employees"
            meta={`${list.length} shown`}
            actions={
              <>
                <Button
                  onClick={() => setShowFilter((v) => !v)}
                  aria-pressed={showFilter}
                  className={cx(showFilter && "border-blue-300 bg-blue-50 text-blue-800")}
                >
                  Filter
                </Button>
                {can.add && (
                  <Button onClick={startNew} disabled={busy}>
                    New
                  </Button>
                )}
              </>
            }
          />
          <div className="space-y-2.5 border-b border-grey-100 px-4 py-2.5">
            <SearchInput
              label="Search employees"
              value={q}
              onChange={setQ}
              placeholder="Search by name or number"
            />
            {showFilter && (
              <div className="grid grid-cols-2 gap-2">
                <Select
                  aria-label="Filter by client"
                  value={clientFilter}
                  onChange={(e) => setClientFilter(e.target.value)}
                >
                  <option value="">All clients</option>
                  {clients.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name}
                    </option>
                  ))}
                </Select>
                <Select
                  aria-label="Filter by payroll"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="">Included and excluded</option>
                  <option value="included">Included in payroll</option>
                  <option value="excluded">Excluded (T, W, X, Y, Z)</option>
                </Select>
              </div>
            )}
          </div>
          <DataTable
            rows={list}
            rowKey={(e) => e.id}
            selected={draft?.id}
            onSelect={open}
            maxHeight="calc(100vh - 300px)"
            rowClass={(e) => isExcluded(e) && "text-grey-500"}
            empty={
              <EmptyState action={can.add ? <Button onClick={startNew}>New</Button> : undefined}>
                {q || clientFilter ? "No employee matches this search." : "No employees yet."}
              </EmptyState>
            }
            columns={[
              {
                key: "no",
                header: "No.",
                width: "56px",
                render: (e) => <span className="font-medium">{e.empNo}</span>,
              },
              {
                key: "name",
                header: "Name",
                render: (e) => (
                  <div className="leading-tight">
                    <div
                      className={cx(
                        "font-medium",
                        isExcluded(e) ? "text-grey-500" : "text-grey-900",
                      )}
                    >
                      {fullName(e)}
                    </div>
                    <div className="text-xs text-grey-500">
                      {e.position} · {db.clients.find((c) => c.id === e.clientId)?.code}
                    </div>
                  </div>
                ),
              },
              {
                key: "ps",
                header: "Payroll",
                render: (e) =>
                  isExcluded(e) ? (
                    <Tag tone="error">{e.payrollStatus} · Excluded</Tag>
                  ) : (
                    <span className="text-grey-700">{e.payrollStatus}</span>
                  ),
              },
              {
                key: "act",
                header: "Active",
                render: (e) => (e.active ? "Yes" : <span className="text-grey-400">No</span>),
              },
            ]}
          />
        </Panel>

        <Panel>
          {!draft ? (
            <>
              <PanelHeader title="No employee open" />
              <div className="space-y-3 px-4 py-4">
                <StatusLine status={status} onDismiss={clear} />
                <p className="text-sm text-grey-600">Choose an employee to see the profile.</p>
              </div>
            </>
          ) : (
            <>
              <PanelHeader
                title={isNew ? "New employee" : `${saved?.empNo} ${saved ? fullName(saved) : ""}`}
                meta={!isNew && isExcluded(draft) ? "Excluded from payroll" : undefined}
                actions={
                  <>
                    {!isNew && can.add && (
                      <Button
                        onClick={() => {
                          setCopyErr(null);
                          setCopyOpen(true);
                        }}
                        disabled={busy}
                      >
                        Copy employee
                      </Button>
                    )}
                    {editable && (
                      <>
                        <Button
                          onClick={() =>
                            isNew ? (setDraft(null), setIsNew(false)) : saved && open(saved)
                          }
                          disabled={busy}
                        >
                          Cancel
                        </Button>
                        <Button variant="primary" onClick={save} busy={busy}>
                          Save
                        </Button>
                      </>
                    )}
                  </>
                }
              />
              {(status || !editable || isExcluded(draft)) && (
                <div className="space-y-2 px-4 pt-4">
                  <StatusLine status={status} onDismiss={clear} />
                  {!editable && (
                    <Notice tone="neutral" title="Read only">
                      Your account cannot edit employees.
                    </Notice>
                  )}
                  {isExcluded(draft) && (
                    <Notice
                      tone="error"
                      title={`Payroll status ${draft.payrollStatus}: excluded from payroll`}
                    >
                      Employees with status {EXCLUDED_STATUSES.join(", ")} are kept here but do not
                      appear in the payroll list.
                    </Notice>
                  )}
                </div>
              )}
              <fieldset disabled={!editable} className="contents">
                <FormSection title="Identity">
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field label="Employee number" required error={err("empNo")}>
                      <TextInput
                        value={draft.empNo}
                        onChange={(e) => set("empNo", e.target.value)}
                      />
                    </Field>
                    <Field label="Internal id">
                      <Readout>{isNew ? "Assigned on save" : draft.id}</Readout>
                    </Field>
                    <Field label="DTR id" hint="Time records join on this." error={err("dtrId")}>
                      <TextInput
                        value={draft.dtrId}
                        onChange={(e) => set("dtrId", e.target.value)}
                      />
                    </Field>
                    <Field label="Gender">
                      <Select value={draft.gender} onChange={(e) => set("gender", e.target.value)}>
                        <option>Female</option>
                        <option>Male</option>
                      </Select>
                    </Field>
                    <Field label="Last name" required error={err("lastName")}>
                      <TextInput
                        value={draft.lastName}
                        onChange={(e) => set("lastName", e.target.value)}
                      />
                    </Field>
                    <Field label="First name" required error={err("firstName")}>
                      <TextInput
                        value={draft.firstName}
                        onChange={(e) => set("firstName", e.target.value)}
                      />
                    </Field>
                    <Field label="Middle name">
                      <TextInput
                        value={draft.middleName}
                        onChange={(e) => set("middleName", e.target.value)}
                      />
                    </Field>
                    <Field label="Civil status">
                      <Select
                        value={draft.civilStatus}
                        onChange={(e) => set("civilStatus", e.target.value)}
                      >
                        {["Single", "Married", "Widowed", "Separated"].map((v) => (
                          <option key={v}>{v}</option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Birth date">
                      <TextInput
                        type="date"
                        value={draft.birthDate}
                        onChange={(e) => set("birthDate", e.target.value)}
                      />
                    </Field>
                    <Field label="Address" className="sm:col-span-2">
                      <TextInput
                        value={draft.address}
                        onChange={(e) => set("address", e.target.value)}
                      />
                    </Field>
                    <Field label="Phone">
                      <TextInput
                        value={draft.phone}
                        onChange={(e) => set("phone", e.target.value)}
                      />
                    </Field>
                  </div>
                </FormSection>

                <FormSection title="Assignment">
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field
                      label="Client"
                      required
                      error={err("clientId")}
                      className="sm:col-span-2"
                    >
                      <Select
                        value={draft.clientId}
                        onChange={(e) => set("clientId", e.target.value)}
                      >
                        {clients.map((c) => (
                          <option key={c.id} value={c.id}>
                            {c.name}
                            {isFmsClient(c) ? " (FMS)" : ""}
                          </option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Department">
                      <Select
                        value={draft.department}
                        onChange={(e) => set("department", e.target.value)}
                      >
                        {db.departments.map((d) => (
                          <option key={d.id}>{d.name}</option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Position">
                      <TextInput
                        value={draft.position}
                        onChange={(e) => set("position", e.target.value)}
                      />
                    </Field>
                    {isFmsClient(client) && (
                      <Field label="Driver or conductor">
                        <Select
                          value={draft.role}
                          onChange={(e) => set("role", e.target.value as Employee["role"])}
                        >
                          <option value="Driver">Driver</option>
                          <option value="Conductor">Conductor</option>
                        </Select>
                      </Field>
                    )}
                    <Field label="Shift">
                      <Select value={draft.shift} onChange={(e) => set("shift", e.target.value)}>
                        {db.shifts.map((s) => (
                          <option key={s.id} value={s.code}>
                            {s.code} ({s.start}–{s.end})
                          </option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Rest day">
                      <Select
                        value={draft.restDay}
                        onChange={(e) => set("restDay", e.target.value)}
                      >
                        <option value="">None</option>
                        {WEEKDAYS.map((d) => (
                          <option key={d}>{d}</option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Payroll status" hint="T, W, X, Y, Z are excluded from payroll.">
                      <Select
                        value={draft.payrollStatus}
                        onChange={(e) =>
                          set("payrollStatus", e.target.value as Employee["payrollStatus"])
                        }
                      >
                        {PAYROLL_STATUSES.map((s) => (
                          <option key={s} value={s}>
                            {s}
                            {EXCLUDED_STATUSES.includes(s) ? " (excluded)" : ""}
                          </option>
                        ))}
                      </Select>
                    </Field>
                    <div className="flex items-end pb-1.5">
                      <Checkbox
                        checked={draft.active}
                        disabled={!editable}
                        onChange={(v) => set("active", v)}
                        label="Active"
                      />
                    </div>
                  </div>
                </FormSection>

                <FormSection
                  title="Pay profile"
                  description="Payroll reads these. The server computes every amount."
                >
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field label="Payment basis">
                      <div
                        role="radiogroup"
                        className="flex h-8 rounded-ctl border border-grey-200 p-0.5"
                      >
                        {(["Monthly", "Hourly"] as const).map((b) => (
                          <button
                            key={b}
                            type="button"
                            role="radio"
                            aria-checked={draft.basis === b}
                            onClick={() => set("basis", b)}
                            className={cx(
                              "flex-1 rounded-[3px] text-sm font-medium",
                              draft.basis === b
                                ? "bg-blue-50 text-blue-800"
                                : "text-grey-600 hover:text-grey-900",
                            )}
                          >
                            {b}
                          </button>
                        ))}
                      </div>
                    </Field>
                    <Field label="Monthly rate">
                      <NumberInput
                        value={draft.monthlyRate}
                        onValue={(n) => set("monthlyRate", n ?? 0)}
                      />
                    </Field>
                    <Field label="Daily rate">
                      <NumberInput
                        value={draft.dailyRate}
                        onValue={(n) => set("dailyRate", n ?? 0)}
                      />
                    </Field>
                    <Field label="Hourly rate" hint="Zero means daily rate / 8.">
                      <NumberInput
                        value={draft.hourlyRate}
                        onValue={(n) => set("hourlyRate", n ?? 0)}
                      />
                    </Field>
                    <Field label="Tax code" error={err("taxCode")}>
                      <Select
                        value={draft.taxCode}
                        onChange={(e) => set("taxCode", e.target.value)}
                      >
                        <option value="">No tax code</option>
                        {db.taxCodes.map((t) => (
                          <option key={t.id} value={t.code}>
                            {t.code} · {t.description}
                          </option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Bank">
                      <Select value={draft.bank} onChange={(e) => set("bank", e.target.value)}>
                        {db.banks.map((b) => (
                          <option key={b.id} value={b.code}>
                            {b.name}
                          </option>
                        ))}
                      </Select>
                    </Field>
                    <Field label="Account number" className="sm:col-span-2">
                      <TextInput
                        value={draft.accountNo}
                        onChange={(e) => set("accountNo", e.target.value)}
                      />
                    </Field>
                  </div>
                </FormSection>

                <FormSection title="Contributions">
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field label="SSS number">
                      <TextInput
                        value={draft.sssNo}
                        onChange={(e) => set("sssNo", e.target.value)}
                      />
                    </Field>
                    <Field label="PhilHealth number">
                      <TextInput
                        value={draft.philhealthNo}
                        onChange={(e) => set("philhealthNo", e.target.value)}
                      />
                    </Field>
                    <Field label="Pag-IBIG number">
                      <TextInput
                        value={draft.pagibigNo}
                        onChange={(e) => set("pagibigNo", e.target.value)}
                      />
                    </Field>
                    <Field label="TIN">
                      <TextInput value={draft.tin} onChange={(e) => set("tin", e.target.value)} />
                    </Field>
                  </div>
                  <div className="mt-3 grid gap-3 sm:grid-cols-4">
                    <div className="space-y-2.5 pt-1 sm:col-span-2">
                      <Checkbox
                        checked={draft.sss}
                        disabled={!editable}
                        onChange={(v) => set("sss", v)}
                        label="Deduct SSS for this person"
                      />
                      <Checkbox
                        checked={draft.philhealth}
                        disabled={!editable}
                        onChange={(v) => set("philhealth", v)}
                        label="Deduct PhilHealth for this person"
                      />
                      <Checkbox
                        checked={draft.pagibig}
                        disabled={!editable}
                        onChange={(v) => set("pagibig", v)}
                        label="Deduct Pag-IBIG for this person"
                      />
                    </div>
                    <Field label="PhilHealth per day" hint="Zero means use the table rule.">
                      <NumberInput
                        value={draft.philhealthPerDay}
                        onValue={(n) => set("philhealthPerDay", n ?? 0)}
                      />
                    </Field>
                    <Field label="Pag-IBIG per day" hint="Zero means use the table rule.">
                      <NumberInput
                        value={draft.pagibigPerDay}
                        onValue={(n) => set("pagibigPerDay", n ?? 0)}
                      />
                    </Field>
                  </div>
                </FormSection>

                <FormSection title="Dates">
                  <div className="grid gap-3 sm:grid-cols-4">
                    <Field label="Hired">
                      <TextInput
                        type="date"
                        value={draft.hired}
                        onChange={(e) => set("hired", e.target.value)}
                      />
                    </Field>
                    <Field label="Contract start">
                      <TextInput
                        type="date"
                        value={draft.contractStart}
                        onChange={(e) => set("contractStart", e.target.value)}
                      />
                    </Field>
                    <Field label="Contract end">
                      <TextInput
                        type="date"
                        value={draft.contractEnd}
                        onChange={(e) => set("contractEnd", e.target.value)}
                      />
                    </Field>
                  </div>
                </FormSection>
              </fieldset>
            </>
          )}
        </Panel>
      </div>

      <Dialog
        open={copyOpen}
        title="Copy employee"
        onClose={() => setCopyOpen(false)}
        footer={
          <>
            <Button onClick={() => setCopyOpen(false)}>Cancel</Button>
            <Button variant="confirm" onClick={doCopy}>
              Copy employee
            </Button>
          </>
        }
      >
        <p className="mb-3">
          Copies the pay profile of <b>{draft ? fullName(draft) : ""}</b> onto a new employee
          number.
        </p>
        <Field label="New employee number" required error={copyErr}>
          <TextInput
            value={copyNo}
            onChange={(e) => setCopyNo(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && doCopy()}
          />
        </Field>
      </Dialog>
    </>
  );
}
