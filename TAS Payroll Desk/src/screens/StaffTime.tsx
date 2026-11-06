import { useState } from "react";
import { useSession } from "../app/session";
import {
  ApiError,
  computeStaff,
  finalizeStaff,
  refreshStaffLeave,
  saveManualTime,
} from "../data/api";
import type { StaffDay } from "../data/types";
import { fullName, leaveLabel } from "../data/rules";
import {
  Button,
  DataTable,
  Dialog,
  EmptyState,
  Field,
  Notice,
  PageHeader,
  Panel,
  PanelHeader,
  StatusLine,
  Tag,
  TextInput,
  cx,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { fmtDateTime, fmtShort, hrs } from "../ui/format";
import { RemarkCell, TimeEmployeeList, TotalsRow, useTimeGuard } from "./TimeShared";

export function StaffTime() {
  const guard = useTimeGuard("staff");
  return (
    <>
      <PageHeader
        section="Time keeping"
        title="Staff time keeping"
        description="Compute the day grid for one employee and half, then finalize it for payroll."
      />
      {guard ?? <StaffTimeBody />}
    </>
  );
}

function StaffTimeBody() {
  const { db, cutoffKey, billed } = useSession();
  const confirm = useConfirm();
  const [empId, setEmpId] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [edit, setEdit] = useState<StaffDay | null>(null);
  const [tin, setTin] = useState("");
  const [tout, setTout] = useState("");
  const [editErr, setEditErr] = useState<{ field?: string; msg: string } | null>(null);
  const { status, setStatus, clear } = useStatus();

  const emp = db.employees.find((e) => e.id === empId);
  const rows = emp ? db.staffFinal[`${emp.id}|${cutoffKey}`] : undefined;
  const fin = emp ? db.payrollAttendance[`${emp.id}|${cutoffKey}`] : undefined;
  const finalized =
    !!fin &&
    !!rows &&
    JSON.stringify(fin.staff) === JSON.stringify(rows.filter((r) => r.remark !== "Invalid Day"));
  const who = emp ? `${emp.empNo} ${fullName(emp)}` : "";

  const run = async (key: string, fn: () => Promise<void>, ok: string, body?: string) => {
    setBusy(key);
    try {
      await fn();
      setStatus({ tone: "success", title: ok, body });
    } catch (e) {
      setStatus({ tone: "error", title: (e as Error).message });
    } finally {
      setBusy(null);
    }
  };
  const onCompute = async () => {
    if (!emp) return;
    if (rows) {
      const ok = await confirm({
        title: "Recompute time records",
        body: (
          <>
            This deletes the current final rows for <b>{who}</b>, key <b>{cutoffKey}</b>, then runs
            the batch again. Manual edits are lost.
          </>
        ),
        confirmLabel: "Recompute",
        tone: "danger",
      });
      if (!ok) return;
    }
    run(
      "compute",
      () => computeStaff(emp.id),
      `Computed time records for ${who}, ${cutoffKey}.`,
      "Check the grid, then Finalize for payroll.",
    );
  };
  const openEdit = (d: StaffDay) => {
    setEdit(d);
    setTin(d.timeIn);
    setTout(d.timeOut);
    setEditErr(null);
  };
  const doEdit = async () => {
    if (!emp || !edit) return;
    try {
      await saveManualTime(emp.id, edit.date, tin, tout);
      setEdit(null);
      setStatus({
        tone: "success",
        title: `Saved manual time for ${who} on ${fmtShort(edit.date)}.`,
        body: "Finalize again so payroll reads the change.",
      });
    } catch (e) {
      setEditErr({ field: e instanceof ApiError ? e.field : undefined, msg: (e as Error).message });
    }
  };

  const sum = (f: (r: StaffDay) => number) => hrs((rows ?? []).reduce((a, r) => a + f(r), 0));
  const primary: "finalize" | null = billed || !rows || finalized ? null : "finalize";

  return (
    <div className="grid items-start gap-4 xl:grid-cols-[240px_minmax(0,1fr)]">
      <TimeEmployeeList
        kind="staff"
        selected={empId}
        onSelect={(e) => {
          setEmpId(e.id);
          clear();
        }}
      />
      <div className="min-w-0 space-y-3">
        <Notice
          tone="info"
          title="Required hours become 8 after Compute. That figure is what payroll pays."
        >
          Work hours, late, and undertime are shown for checking. A manual edit changes one punch
          pair only; it is not the pay rule.
        </Notice>
        <Panel>
          {!emp ? (
            <>
              <PanelHeader title="No employee open" />
              <EmptyState>
                Choose an employee on the left to open the day grid for {cutoffKey}.
              </EmptyState>
            </>
          ) : (
            <>
              <PanelHeader
                title={who}
                meta={
                  billed
                    ? "Billed cutoff, read-only"
                    : finalized
                      ? `Finalized ${fmtDateTime(fin!.at)}`
                      : fin
                        ? "Changed since finalize"
                        : rows
                          ? "Computed, not finalized"
                          : "Not computed"
                }
                actions={
                  !billed && (
                    <>
                      <Button
                        onClick={() =>
                          run(
                            "refresh",
                            () => refreshStaffLeave(emp.id),
                            `Refreshed leave and overtime for ${who}.`,
                            "Finalize again so payroll reads the change.",
                          )
                        }
                        disabled={!rows || !!busy}
                        busy={busy === "refresh"}
                      >
                        Refresh from leave and overtime
                      </Button>
                      {rows && (
                        <Button onClick={onCompute} busy={busy === "compute"} disabled={!!busy}>
                          Compute
                        </Button>
                      )}
                      <Button
                        variant={primary === "finalize" ? "primary" : "secondary"}
                        onClick={() =>
                          run(
                            "finalize",
                            () => finalizeStaff(emp.id),
                            `Finalized ${who} for payroll, ${cutoffKey}.`,
                            "Payroll entry can now load this employee.",
                          )
                        }
                        disabled={!rows || finalized || !!busy}
                        busy={busy === "finalize"}
                      >
                        Finalize for payroll
                      </Button>
                    </>
                  )
                }
              />
              {rows && !billed && (
                <p className="border-b border-grey-100 px-4 py-2 text-xs text-grey-500">
                  Choose a day to make a manual edit of its punch pair.
                </p>
              )}
              {status && (
                <div className="border-b border-grey-100 px-4 py-3">
                  <StatusLine status={status} onDismiss={clear} />
                </div>
              )}
              {!rows ? (
                <EmptyState
                  action={
                    !billed && (
                      <Button variant="primary" onClick={onCompute} busy={busy === "compute"}>
                        Compute
                      </Button>
                    )
                  }
                >
                  {billed
                    ? "This cutoff is billed. Working time rows were removed when payroll moved to history."
                    : `No computed rows for ${cutoffKey}. If finalize is skipped, payroll says no time records were processed.`}
                </EmptyState>
              ) : (
                <DataTable
                  rows={rows}
                  rowKey={(r) => r.date}
                  onSelect={billed ? undefined : (r) => r.remark !== "Invalid Day" && openEdit(r)}
                  dense
                  maxHeight="calc(100vh - 330px)"
                  rowClass={(r) =>
                    cx(
                      (r.remark === "Day off" || r.remark === "Invalid Day") && "text-grey-400",
                      (r.remark === "Absent" || r.remark === "No In/Out") && "bg-red-50/60",
                    )
                  }
                  columns={[
                    {
                      key: "d",
                      header: "Date",
                      render: (r) => (
                        <span>
                          <span className="font-medium">{fmtShort(r.date)}</span>{" "}
                          <span className="text-grey-500">{r.weekday}</span>
                        </span>
                      ),
                    },
                    {
                      key: "in",
                      header: "Time in",
                      render: (r) => r.timeIn || <span className="text-grey-300">—</span>,
                    },
                    {
                      key: "out",
                      header: "Time out",
                      render: (r) => r.timeOut || <span className="text-grey-300">—</span>,
                    },
                    {
                      key: "req",
                      header: "Required",
                      align: "right",
                      render: (r) => (
                        <span
                          className={cx(
                            "font-semibold",
                            r.required ? "text-grey-900" : "text-grey-300",
                          )}
                        >
                          {hrs(r.required)}
                        </span>
                      ),
                    },
                    {
                      key: "work",
                      header: "Work",
                      align: "right",
                      render: (r) => <Num v={r.work} />,
                    },
                    {
                      key: "late",
                      header: "Late",
                      align: "right",
                      render: (r) => <Num v={r.late} bad />,
                    },
                    {
                      key: "ut",
                      header: "Undertime",
                      align: "right",
                      render: (r) => <Num v={r.undertime} bad />,
                    },
                    {
                      key: "ot",
                      header: "Overtime",
                      align: "right",
                      render: (r) => <Num v={r.overtime} />,
                    },
                    {
                      key: "nd",
                      header: "Night diff",
                      align: "right",
                      render: (r) => <Num v={r.nd} />,
                    },
                    {
                      key: "lh",
                      header: "Leave hrs",
                      align: "right",
                      render: (r) => <Num v={r.leaveHours} />,
                    },
                    {
                      key: "lv",
                      header: "Leave",
                      render: (r) =>
                        r.leaveLetter ? (
                          <span>
                            <b className="font-semibold">{r.leaveLetter}</b>{" "}
                            <span className="text-grey-500">{leaveLabel(r.leaveLetter)}</span>
                          </span>
                        ) : null,
                    },
                    {
                      key: "rm",
                      header: "Remark",
                      render: (r) => (
                        <span className="inline-flex items-center gap-1.5">
                          <RemarkCell remark={r.remark} />
                          {r.manual && <Tag tone="blue">Manual edit</Tag>}
                        </span>
                      ),
                    },
                  ]}
                  foot={
                    <TotalsRow
                      cells={[
                        "Totals",
                        null,
                        null,
                        sum((r) => r.required),
                        sum((r) => r.work),
                        sum((r) => r.late),
                        sum((r) => r.undertime),
                        sum((r) => r.overtime),
                        sum((r) => r.nd),
                        sum((r) => r.leaveHours),
                        null,
                        null,
                      ]}
                    />
                  }
                />
              )}
            </>
          )}
        </Panel>
      </div>

      <Dialog
        open={!!edit}
        title={`Manual edit, ${edit ? fmtShort(edit.date) : ""}`}
        onClose={() => setEdit(null)}
        footer={
          <>
            <Button onClick={() => setEdit(null)}>Cancel</Button>
            <Button variant="confirm" onClick={doEdit}>
              Save manual time
            </Button>
          </>
        }
      >
        <p className="mb-3">
          Edits one punch pair for {who}. Required hours stay at 8 for a processed day; this is a
          correction, not a pay rule.
        </p>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Time in" error={editErr?.field === "in" ? editErr.msg : null}>
            <TextInput type="time" value={tin} onChange={(e) => setTin(e.target.value)} />
          </Field>
          <Field label="Time out" error={editErr?.field === "out" ? editErr.msg : null}>
            <TextInput type="time" value={tout} onChange={(e) => setTout(e.target.value)} />
          </Field>
        </div>
        {editErr && !editErr.field && (
          <p className="mt-2 text-sm font-medium text-red-600">{editErr.msg}</p>
        )}
      </Dialog>
    </div>
  );
}

function Num({ v, bad }: { v: number; bad?: boolean }) {
  if (!v) return <span className="text-grey-300">0.00</span>;
  return <span className={bad ? "font-medium text-red-600" : "text-grey-900"}>{hrs(v)}</span>;
}
