import { useEffect, useMemo, useState, type ReactNode } from "react";
import { useSession } from "../app/session";
import { deleteRow, saveRow } from "../data/api";
import {
  Button,
  Checkbox,
  DataTable,
  EmptyState,
  Field,
  Notice,
  NumberInput,
  Panel,
  PanelHeader,
  Select,
  StatusLine,
  TextInput,
  cx,
  useConfirm,
  useStatus,
  type Col,
} from "./kit";
import { uid } from "../data/money";

export type FieldDef<T> = {
  key: keyof T & string;
  label: string;
  type: "text" | "number" | "date" | "time" | "select" | "checkbox";
  options?: { value: string; label: string }[];
  required?: boolean;
  hint?: string;
  span?: 1 | 2;
  decimals?: number;
};

type Table = Parameters<typeof saveRow>[0];

export function RecordEditor<T extends { id: string }>({
  title,
  noun,
  table,
  rows,
  columns,
  fields,
  blank,
  emptyText,
  describe,
  toolbar,
  disabledReason,
  maxHeight = "calc(100vh - 280px)",
  formNote,
}: {
  title: string;
  noun: string;
  table: Table;
  rows: T[];
  columns: Col<T>[];
  fields: FieldDef<T>[];
  blank: () => Omit<T, "id">;
  emptyText: string;
  describe: (row: T) => string;
  toolbar?: ReactNode;
  disabledReason?: string | null;
  maxHeight?: string;
  formNote?: ReactNode;
}) {
  const { can } = useSession();
  const confirm = useConfirm();
  const { status, setStatus, clear } = useStatus();
  const [selId, setSelId] = useState<string | null>(null);
  const [draft, setDraft] = useState<T | null>(null);
  const [isNew, setIsNew] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [busy, setBusy] = useState<"save" | "delete" | null>(null);

  const saved = useMemo(() => rows.find((r) => r.id === selId) ?? null, [rows, selId]);
  useEffect(() => {
    if (!isNew && selId && !saved) {
      setSelId(null);
      setDraft(null);
    }
  }, [saved, selId, isNew]);

  const editable = !disabledReason && (isNew ? can.add : can.edit);
  const dirty = draft && JSON.stringify(draft) !== JSON.stringify(isNew ? null : saved);

  const open = (r: T) => {
    setSelId(r.id);
    setDraft({ ...r });
    setIsNew(false);
    setErrors({});
    clear();
  };
  const startNew = () => {
    const row = { ...blank(), id: uid() } as T;
    setSelId(row.id);
    setDraft(row);
    setIsNew(true);
    setErrors({});
    clear();
  };
  const cancel = () => {
    setErrors({});
    clear();
    if (isNew) {
      setSelId(null);
      setDraft(null);
      setIsNew(false);
    } else if (saved) setDraft({ ...saved });
  };
  const save = async () => {
    if (!draft) return;
    const errs: Record<string, string> = {};
    for (const f of fields) {
      const v = draft[f.key] as unknown;
      if (f.required && (v === "" || v == null)) errs[f.key] = `${f.label} is required.`;
    }
    setErrors(errs);
    if (Object.keys(errs).length) {
      setStatus({
        tone: "error",
        title: `${noun[0].toUpperCase()}${noun.slice(1)} not saved`,
        body: Object.values(errs)[0],
      });
      return;
    }
    setBusy("save");
    await saveRow(table, draft, noun);
    setBusy(null);
    setIsNew(false);
    setStatus({ tone: "success", title: `Saved ${noun} ${describe(draft)}.` });
  };
  const remove = async () => {
    if (!saved) return;
    const ok = await confirm({
      title: `Delete ${noun}`,
      body: (
        <>
          Delete {noun} <b>{describe(saved)}</b>? This cannot be undone.
        </>
      ),
      confirmLabel: "Delete",
      tone: "danger",
    });
    if (!ok) return;
    setBusy("delete");
    await deleteRow(table, saved.id, noun);
    setBusy(null);
    setSelId(null);
    setDraft(null);
    setStatus({ tone: "success", title: `Deleted ${noun} ${describe(saved)}.` });
  };

  const set = <K extends keyof T>(k: K, v: T[K]) => setDraft((d) => (d ? { ...d, [k]: v } : d));
  const checks = fields.filter((f) => f.type === "checkbox");
  const inputs = fields.filter((f) => f.type !== "checkbox");

  return (
    <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1.2fr)_minmax(400px,1fr)]">
      <Panel>
        <PanelHeader
          title={title}
          meta={`${rows.length} ${rows.length === 1 ? "row" : "rows"}`}
          actions={
            can.add && !disabledReason ? (
              <Button onClick={startNew} disabled={busy !== null}>
                New
              </Button>
            ) : null
          }
        />
        {toolbar && (
          <div className="flex flex-wrap items-center gap-2 border-b border-grey-100 px-4 py-2.5">
            {toolbar}
          </div>
        )}
        <DataTable
          columns={columns}
          rows={rows}
          rowKey={(r) => r.id}
          selected={selId}
          onSelect={open}
          maxHeight={maxHeight}
          empty={
            <EmptyState
              action={
                can.add && !disabledReason ? <Button onClick={startNew}>New</Button> : undefined
              }
            >
              {emptyText}
            </EmptyState>
          }
        />
      </Panel>

      <Panel className="xl:sticky xl:top-[7.5rem]">
        {draft ? (
          <>
            <PanelHeader
              title={isNew ? `New ${noun}` : describe(saved ?? draft)}
              meta={dirty ? "Unsaved changes" : undefined}
              actions={
                <>
                  {!isNew && can.del && !disabledReason && (
                    <Button
                      variant="danger"
                      onClick={remove}
                      busy={busy === "delete"}
                      disabled={busy !== null}
                    >
                      Delete
                    </Button>
                  )}
                  {editable && (
                    <>
                      <Button onClick={cancel} disabled={busy !== null}>
                        Cancel
                      </Button>
                      <Button
                        variant="primary"
                        onClick={save}
                        busy={busy === "save"}
                        disabled={busy !== null}
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
              {!editable && (
                <Notice tone="neutral" title="Read only">
                  {disabledReason ?? "Your account cannot edit this table."}
                </Notice>
              )}
              {formNote}
              <div className="grid grid-cols-2 gap-x-3 gap-y-3">
                {inputs.map((f) => (
                  <Field
                    key={f.key}
                    label={f.label}
                    required={f.required}
                    hint={f.hint}
                    error={errors[f.key]}
                    className={cx(
                      f.span !== 1 && "col-span-2",
                      f.span === 1 && "col-span-2 sm:col-span-1",
                    )}
                  >
                    {f.type === "select" ? (
                      <Select
                        value={String(draft[f.key] ?? "")}
                        disabled={!editable}
                        onChange={(e) => set(f.key, e.target.value as T[typeof f.key])}
                      >
                        <option value="">Choose</option>
                        {f.options?.map((o) => (
                          <option key={o.value} value={o.value}>
                            {o.label}
                          </option>
                        ))}
                      </Select>
                    ) : f.type === "number" ? (
                      <NumberInput
                        value={(draft[f.key] as number | null) ?? null}
                        decimals={f.decimals ?? 2}
                        disabled={!editable}
                        onValue={(n) => set(f.key, (n ?? 0) as T[typeof f.key])}
                      />
                    ) : (
                      <TextInput
                        type={f.type}
                        value={String(draft[f.key] ?? "")}
                        disabled={!editable}
                        onChange={(e) => set(f.key, e.target.value as T[typeof f.key])}
                      />
                    )}
                  </Field>
                ))}
              </div>
              {checks.length > 0 && (
                <div className="grid grid-cols-1 gap-2.5 border-t border-grey-100 pt-4 sm:grid-cols-2">
                  {checks.map((f) => (
                    <Checkbox
                      key={f.key}
                      checked={!!draft[f.key]}
                      disabled={!editable}
                      hint={f.hint}
                      onChange={(v) => set(f.key, v as T[typeof f.key])}
                      label={f.label}
                    />
                  ))}
                </div>
              )}
            </div>
          </>
        ) : (
          <>
            <PanelHeader title={`No ${noun} open`} />
            <div className="space-y-3 px-4 py-4">
              <StatusLine status={status} onDismiss={clear} />
              <p className="text-sm text-grey-600">
                Choose a row on the left to edit it
                {can.add && !disabledReason ? ", or use New to add one" : ""}.
              </p>
            </div>
          </>
        )}
      </Panel>
    </div>
  );
}
