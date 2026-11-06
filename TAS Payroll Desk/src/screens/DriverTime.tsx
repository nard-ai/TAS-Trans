import { useState } from "react";
import { useSession } from "../app/session";
import { computeFms, finalizeFms, refreshFmsLeave } from "../data/api";
import type { FmsDay } from "../data/types";
import { fullName, leaveLabel } from "../data/rules";
import {
  Button,
  DataTable,
  EmptyState,
  Notice,
  PageHeader,
  Panel,
  PanelHeader,
  StatusLine,
  cx,
  useConfirm,
  useStatus,
} from "../ui/kit";
import { fmtDateTime, hrs } from "../ui/format";
import { RemarkCell, TimeEmployeeList, TotalsRow, useTimeGuard } from "./TimeShared";

export function DriverTime() {
  const guard = useTimeGuard("fms");
  return (
    <>
      <PageHeader
        section="Time keeping"
        title="Driver time keeping"
        description="Up to six in and out pairs per day. Required hours are the sum of the pairs."
      />
      {guard ?? <DriverTimeBody />}
    </>
  );
}

function DriverTimeBody() {
  const { db, cutoffKey, billed } = useSession();
  const confirm = useConfirm();
  const [empId, setEmpId] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const { status, setStatus, clear } = useStatus();

  const emp = db.employees.find((e) => e.id === empId);
  const rows = emp ? db.fmsFinal[`${emp.id}|${cutoffKey}`] : undefined;
  const fin = emp ? db.payrollAttendance[`${emp.id}|${cutoffKey}`] : undefined;
  const finalized = !!fin && !!rows && JSON.stringify(fin.fms) === JSON.stringify(rows);
  const who = emp ? `${emp.empNo} ${fullName(emp)}` : "";
  const primary: "finalize" | null = billed || !rows || finalized ? null : "finalize";

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
            the batch again.
          </>
        ),
        confirmLabel: "Recompute",
        tone: "danger",
      });
      if (!ok) return;
    }
    run(
      "compute",
      () => computeFms(emp.id),
      `Computed driver time records for ${who}, ${cutoffKey}.`,
      "Check the grid, then Finalize for payroll.",
    );
  };

  const pairCell = (r: FmsDay, i: number) => {
    const p = r.pairs[i];
    if (!p) return <span className="text-grey-300">—</span>;
    return (
      <div className="leading-tight">
        <div className="whitespace-nowrap text-grey-900">
          {p.in}–{p.out}
        </div>
        <div className="text-xs text-grey-500">{hrs(r.pairHours[i])} h</div>
      </div>
    );
  };
  const blank = <span className="text-grey-300">—</span>;
  // Show only as many pair columns as this person used (never more than six).
  const pairIdx = Array.from(
    { length: Math.min(6, Math.max(1, ...(rows ?? []).map((r) => r.pairs.length))) },
    (_, i) => i,
  );
  const sum = (f: (r: FmsDay) => number) => hrs((rows ?? []).reduce((a, r) => a + f(r), 0));

  return (
    <div className="grid items-start gap-4 xl:grid-cols-[240px_minmax(0,1fr)]">
      <TimeEmployeeList
        kind="fms"
        selected={empId}
        onSelect={(e) => {
          setEmpId(e.id);
          clear();
        }}
      />
      <div className="min-w-0 space-y-3">
        <Notice tone="info" title="Overtime from this compute is 0.">
          Rest-day overtime is applied on the last day only, and only for leave letter D. Night
          differential and late stay blank unless a later leave update sets them.
        </Notice>
        <Panel>
          {!emp ? (
            <>
              <PanelHeader title="No driver or conductor open" />
              <EmptyState>
                Choose a driver or conductor on the left to open the day grid for {cutoffKey}.
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
                            () => refreshFmsLeave(emp.id),
                            `Refreshed leave and overtime for ${who}.`,
                            "Finalize again so payroll reads the change.",
                          )
                        }
                        disabled={!rows || !!busy}
                        busy={busy === "refresh"}
                      >
                        Refresh leave and overtime
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
                            () => finalizeFms(emp.id),
                            `Finalized ${who} for payroll, ${cutoffKey}.`,
                            "Payroll entry can now load this driver.",
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
                    : `No computed rows for ${cutoffKey}.`}
                </EmptyState>
              ) : (
                <DataTable
                  rows={rows}
                  rowKey={(r) => String(r.day)}
                  dense
                  maxHeight="calc(100vh - 330px)"
                  rowClass={(r) =>
                    cx(
                      r.remark === "Invalid Day" && "bg-grey-50 text-grey-400",
                      (r.remark === "Absent" || r.remark === "No In/Out") && "bg-red-50/60",
                    )
                  }
                  columns={[
                    {
                      key: "d",
                      header: "Day",
                      render: (r) => <span className="font-medium">{r.day}</span>,
                    },
                    ...pairIdx.map((i) => ({
                      key: `p${i}`,
                      header: `Pair ${i + 1}`,
                      render: (r: FmsDay) => pairCell(r, i),
                    })),
                    {
                      key: "req",
                      header: "Required",
                      align: "right" as const,
                      render: (r: FmsDay) => (
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
                      key: "nd",
                      header: "Night diff",
                      align: "right" as const,
                      render: (r: FmsDay) => (r.nd == null ? blank : hrs(r.nd)),
                    },
                    {
                      key: "late",
                      header: "Late",
                      align: "right" as const,
                      render: (r: FmsDay) => (r.late == null ? blank : hrs(r.late)),
                    },
                    {
                      key: "lv",
                      header: "Leave",
                      render: (r: FmsDay) =>
                        r.leaveLetter ? (
                          <div className="leading-tight">
                            <b className="font-semibold">{r.leaveLetter}</b>{" "}
                            <span className="text-grey-500">{hrs(r.leaveHours)} h</span>
                            <div className="text-xs text-grey-500">{leaveLabel(r.leaveLetter)}</div>
                          </div>
                        ) : null,
                    },
                    {
                      key: "rd",
                      header: "Rest-day OT",
                      align: "right" as const,
                      render: (r: FmsDay) =>
                        r.restDayOt ? (
                          <span className="font-medium">{hrs(r.restDayOt)}</span>
                        ) : (
                          blank
                        ),
                    },
                    {
                      key: "rm",
                      header: "Remark",
                      render: (r: FmsDay) => <RemarkCell remark={r.remark} />,
                    },
                  ]}
                  foot={
                    <TotalsRow
                      cells={[
                        "Totals",
                        ...pairIdx.map(() => null),
                        sum((r) => r.required),
                        null,
                        null,
                        null,
                        sum((r) => r.restDayOt),
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
    </div>
  );
}
