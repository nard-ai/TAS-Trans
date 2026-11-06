import { useRef, useState } from "react";
import { useSession } from "../app/session";
import {
  ApiError,
  importPunches,
  previewPunches,
  replaceDtrId,
  type PreviewPunch,
} from "../data/api";
import type { Punch } from "../data/types";
import {
  Button,
  DataTable,
  EmptyState,
  Field,
  Notice,
  PageHeader,
  Panel,
  PanelHeader,
  Readout,
  Select,
  StatusLine,
  Tag,
  TextInput,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { fmtDate } from "../ui/format";

const SAMPLE = `D1009,2025-11-17,07:52,In
D1009,2025-11-17,17:06,Out
D1009,2025-11-18,07:49,In
D1009,2025-11-18,17:02,Out
D1012,2025-11-18,07:58,In
D1003,2025-11-18,25:10,Out`;

function parse(text: string, source: string): Omit<Punch, "id">[] {
  return text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => {
      const [dtrId = "", date = "", time = "", dir = ""] = l.split(/[,\t;]/).map((x) => x.trim());
      const d = dir.toLowerCase();
      return {
        dtrId,
        date,
        time,
        dir: (d === "in" ? "In" : d === "out" ? "Out" : dir) as Punch["dir"],
        source,
      };
    });
}

export function ImportDtr() {
  const { db, user, cutoff, cutoffClient } = useSession();
  const confirm = useConfirm();
  const fileRef = useRef<HTMLInputElement>(null);
  const [source, setSource] = useState<"file" | "device">("file");
  const [fileName, setFileName] = useState("");
  const [text, setText] = useState("");
  const [preview, setPreview] = useState<PreviewPunch[] | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [oldId, setOldId] = useState("");
  const [newId, setNewId] = useState("");
  const [repField, setRepField] = useState<string | undefined>();
  const { status, setStatus, clear } = useStatus();
  const rep = useStatus();

  const inCutoff = cutoff
    ? db.punches.filter(
        (p) =>
          p.date >= cutoff.start &&
          p.date <= cutoff.end &&
          db.employees.some((e) => e.dtrId === p.dtrId && e.clientId === cutoff.clientId),
      )
    : [];
  const bad = preview?.filter((r) => r.problem).length ?? 0;
  const good = (preview?.length ?? 0) - bad;

  const onFile = async (f: File | undefined) => {
    if (!f) return;
    setFileName(f.name);
    setText(await f.text());
    setPreview(null);
  };
  const doPreview = async () => {
    clear();
    const raw = source === "device" ? SAMPLE : text;
    if (!raw.trim()) {
      setStatus({
        tone: "error",
        title: "Nothing to preview",
        body: "Choose a punch file first. Each line: DTR id, date, time, In or Out.",
      });
      return;
    }
    setBusy("preview");
    setPreview(
      await previewPunches(
        parse(raw, source === "device" ? "Terminal TDX-01" : fileName || "File"),
      ),
    );
    setBusy(null);
  };
  const doImport = async () => {
    if (!preview) return;
    setBusy("import");
    try {
      const r = await importPunches(preview);
      setStatus({
        tone: "success",
        title: `Imported ${r.imported} punches.`,
        body: r.skipped ? `${r.skipped} rows with problems were skipped.` : undefined,
      });
      setPreview(null);
      setText("");
      setFileName("");
    } catch (e) {
      setStatus({ tone: "error", title: "Import failed", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const doReplace = async () => {
    rep.clear();
    setRepField(undefined);
    const count = db.punches.filter((p) => p.dtrId === oldId.trim()).length;
    const ok = await confirm({
      title: "Replace DTR id",
      body: (
        <>
          Move <b>{count}</b> punches from DTR id <b>{oldId || "(blank)"}</b> to{" "}
          <b>{newId || "(blank)"}</b>? If the new id cannot be saved, the old rows stay as they are.
        </>
      ),
      confirmLabel: "Replace DTR id",
      tone: "danger",
    });
    if (!ok) return;
    setBusy("replace");
    try {
      const r = await replaceDtrId(oldId.trim(), newId.trim());
      rep.setStatus({
        tone: "success",
        title: `Replaced DTR id ${oldId} with ${newId} on ${r.count} punches.`,
        body: `Now joined to ${r.owner}.`,
      });
      setOldId("");
      setNewId("");
    } catch (e) {
      setRepField(e instanceof ApiError ? e.field : undefined);
      rep.setStatus({ tone: "error", title: "DTR id not replaced", body: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };

  return (
    <>
      <PageHeader
        section="Time keeping"
        title="Import time records"
        description="Bring punches in before the time-record batch runs. Preview first; only rows without problems are saved."
      />
      <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1.4fr)_minmax(320px,1fr)]">
        <div className="min-w-0 space-y-4">
          <Panel>
            <PanelHeader
              title="Punches"
              actions={
                <>
                  <Button onClick={doPreview} busy={busy === "preview"} disabled={!!busy}>
                    Preview
                  </Button>
                  <Button
                    variant="primary"
                    onClick={doImport}
                    busy={busy === "import"}
                    disabled={!preview || good === 0 || !!busy}
                  >
                    Import{preview ? ` ${good} rows` : ""}
                  </Button>
                </>
              }
            />
            <div className="space-y-3 border-b border-grey-100 px-4 py-3">
              <StatusLine status={status} onDismiss={clear} />
              <div className="grid gap-3 sm:grid-cols-[200px_1fr]">
                <Field label="Source">
                  <Select
                    value={source}
                    onChange={(e) => {
                      setSource(e.target.value as "file" | "device");
                      setPreview(null);
                    }}
                  >
                    <option value="file">Punch file</option>
                    <option value="device">Terminal TDX-01</option>
                  </Select>
                </Field>
                {source === "file" ? (
                  <Field
                    label="File"
                    hint="CSV or text. Each line: DTR id, date (YYYY-MM-DD), time (HH:MM), In or Out."
                  >
                    <div className="flex gap-2">
                      <input
                        ref={fileRef}
                        type="file"
                        accept=".csv,.txt"
                        className="sr-only"
                        onChange={(e) => onFile(e.target.files?.[0])}
                      />
                      <Button onClick={() => fileRef.current?.click()}>Choose file</Button>
                      <Readout>
                        {fileName || <span className="text-grey-400">No file chosen</span>}
                      </Readout>
                    </div>
                  </Field>
                ) : (
                  <Field label="Terminal">
                    <Readout>Last sync Dec 2, 2025 06:00 · punches since the last import</Readout>
                  </Field>
                )}
              </div>
            </div>
            {preview ? (
              <>
                <div className="flex items-center gap-3 border-b border-grey-100 px-4 py-2 text-sm">
                  <span className="font-medium text-grey-900">Preview</span>
                  <span className="text-cyan-800">{good} ready</span>
                  {bad > 0 && (
                    <span className="font-medium text-red-600">
                      {bad} with problems, will be skipped
                    </span>
                  )}
                </div>
                <DataTable
                  rows={preview}
                  rowKey={(r) => `${r.dtrId}${r.date}${r.time}${r.dir}`}
                  rowClass={(r) => !!r.problem && "bg-red-50"}
                  maxHeight="420px"
                  columns={[
                    {
                      key: "id",
                      header: "DTR id",
                      render: (r) => <span className="font-medium">{r.dtrId}</span>,
                    },
                    {
                      key: "emp",
                      header: "Employee",
                      render: (r) => r.employee ?? <span className="text-grey-400">—</span>,
                    },
                    { key: "d", header: "Date", render: (r) => r.date },
                    { key: "t", header: "Time", render: (r) => r.time },
                    { key: "io", header: "In or out", render: (r) => r.dir },
                    {
                      key: "p",
                      header: "Check",
                      render: (r) =>
                        r.problem ? (
                          <span className="text-xs font-medium text-red-600">{r.problem}</span>
                        ) : (
                          <span className="text-xs text-cyan-800">Ready</span>
                        ),
                    },
                  ]}
                />
              </>
            ) : (
              <>
                <div className="flex items-center justify-between border-b border-grey-100 px-4 py-2 text-sm">
                  <span className="font-medium text-grey-900">
                    Saved punches in the selected cutoff
                  </span>
                  <span className="text-grey-500">
                    {cutoffClient
                      ? `${cutoffClient.name}, ${cutoff?.start} to ${cutoff?.end}`
                      : "No cutoff selected"}
                  </span>
                </div>
                <DataTable
                  rows={inCutoff.slice(-40).reverse()}
                  rowKey={(p) => p.id}
                  maxHeight="420px"
                  empty={
                    <EmptyState>
                      No punches for this cutoff yet. Choose a file or the terminal, then Preview.
                    </EmptyState>
                  }
                  columns={[
                    {
                      key: "id",
                      header: "DTR id",
                      render: (p) => <span className="font-medium">{p.dtrId}</span>,
                    },
                    { key: "d", header: "Date", render: (p) => fmtDate(p.date) },
                    { key: "t", header: "Time", render: (p) => p.time },
                    {
                      key: "io",
                      header: "In or out",
                      render: (p) => <Tag tone={p.dir === "In" ? "blue" : "neutral"}>{p.dir}</Tag>,
                    },
                    {
                      key: "s",
                      header: "Source",
                      render: (p) => <span className="text-grey-500">{p.source}</span>,
                    },
                  ]}
                />
                {inCutoff.length > 40 && (
                  <p className="px-4 py-2 text-xs text-grey-500">
                    Showing the latest 40 of {inCutoff.length} punches.
                  </p>
                )}
              </>
            )}
          </Panel>
        </div>

        <Panel>
          <PanelHeader
            title="Replace DTR id"
            actions={
              <Button
                variant="danger"
                onClick={doReplace}
                busy={busy === "replace"}
                disabled={!!busy}
              >
                Replace DTR id
              </Button>
            }
          />
          <div className="space-y-3 px-4 py-4">
            <p className="text-sm text-grey-600">
              Moves punches saved under an invalid DTR id onto the right one.
            </p>
            <StatusLine status={rep.status} onDismiss={rep.clear} />
            <Field
              label="Old invalid DTR id"
              required
              error={repField === "old" ? (rep.status?.body as string) : null}
            >
              <TextInput value={oldId} onChange={(e) => setOldId(e.target.value)} />
            </Field>
            <Field
              label="New DTR id"
              required
              error={repField === "new" ? (rep.status?.body as string) : null}
            >
              <TextInput value={newId} onChange={(e) => setNewId(e.target.value)} />
            </Field>
            <Field label="Changed by">
              <Readout>{user.username}</Readout>
            </Field>
            <Notice tone="neutral" title="Confirmation required">
              Replace asks before it runs. If the new id cannot be saved, the old rows stay.
            </Notice>
          </div>
        </Panel>
      </div>
    </>
  );
}
