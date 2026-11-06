import { useSession } from "../app/session";
import { navigate } from "../app/router";
import { selectCutoff } from "../data/api";
import {
  canRunPayroll,
  canSee,
  cutoffHalf,
  fullName,
  isExcluded,
  isFmsClient,
} from "../data/rules";
import { Button, DataTable, EmptyState, PageHeader, Panel, PanelHeader, Tag, cx } from "../ui/kit";
import { Icon } from "../ui/icons";
import { fmtDate } from "../ui/format";
import type { CutoffRecord } from "../data/types";

export function Home() {
  const s = useSession();
  const { db, user, cutoff, cutoffClient, cutoffKey, fms, billed } = s;
  const visibleIds = new Set(s.clients.map((c) => c.id));
  const open = db.cutoffs
    .filter((c) => !c.billed && visibleIds.has(c.clientId))
    .sort((a, b) => b.end.localeCompare(a.end));

  const counts = (c: CutoffRecord) => {
    const emps = db.employees.filter((e) => e.clientId === c.clientId && !isExcluded(e));
    const processed = emps.filter((e) =>
      db.openPayroll.some((p) => p.empId === e.id && p.key === c.key && p.clientId === c.clientId),
    ).length;
    return { total: emps.length, processed, remaining: emps.length - processed };
  };

  const queue =
    cutoff && cutoffClient
      ? db.employees
          .filter((e) => e.clientId === cutoffClient.id && !isExcluded(e))
          .filter(
            (e) =>
              !db.openPayroll.some(
                (p) => p.empId === e.id && p.key === cutoffKey && p.clientId === cutoffClient.id,
              ),
          )
          .map((e) => ({
            e,
            computed: !!(fms ? db.fmsFinal : db.staffFinal)[`${e.id}|${cutoffKey}`],
            finalized: !!db.payrollAttendance[`${e.id}|${cutoffKey}`],
          }))
      : [];

  const payroll = canRunPayroll(user);
  const time = canSee(user, "timeKeeping");
  const shortcuts = [
    time && {
      label: fms ? "Driver time keeping" : "Staff time keeping",
      note: "Compute and finalize attendance for the selected cutoff.",
      to: fms ? "time-fms" : "time-staff",
    },
    payroll && {
      label: "Payroll entry",
      note: "Load attendance, check amounts, save each employee.",
      to: "payroll-entry",
      primary: true,
    },
    canSee(user, "payroll") &&
      s.clients.some(isFmsClient) && {
        label: "Driver commissions",
        note: "Load FMS trips and save a monthly payment.",
        to: "commissions",
      },
  ].filter(Boolean) as { label: string; note: string; to: string; primary?: boolean }[];

  return (
    <>
      <PageHeader
        title="Work queue"
        description="Open cutoffs for the clients you can open, and who is still waiting for payroll."
      />

      <div className="grid items-start gap-4 xl:grid-cols-[minmax(0,1fr)_340px]">
        <div className="min-w-0 space-y-4">
          <Panel>
            <PanelHeader title="Open cutoffs" meta={`${open.length} open`} />
            <DataTable
              rows={open}
              rowKey={(r) => `${r.clientId}|${r.key}`}
              selected={cutoff ? `${cutoff.clientId}|${cutoffKey}` : null}
              empty={
                <EmptyState
                  action={<Button onClick={() => navigate("cutoff")}>Select cutoff</Button>}
                >
                  No open cutoffs for your clients.
                </EmptyState>
              }
              columns={[
                {
                  key: "client",
                  header: "Client",
                  render: (r) => (
                    <span className="font-medium">
                      {db.clients.find((c) => c.id === r.clientId)?.name}
                    </span>
                  ),
                },
                {
                  key: "type",
                  header: "Type",
                  render: (r) =>
                    isFmsClient(db.clients.find((c) => c.id === r.clientId)) ? (
                      <Tag tone="blue">FMS</Tag>
                    ) : (
                      <Tag>Staff</Tag>
                    ),
                },
                { key: "start", header: "Start", render: (r) => fmtDate(r.start) },
                { key: "end", header: "End", render: (r) => fmtDate(r.end) },
                { key: "half", header: "Half", align: "right", render: (r) => cutoffHalf(r.end) },
                { key: "key", header: "DTR key", render: (r) => r.key },
                {
                  key: "status",
                  header: "Status",
                  render: (r) =>
                    r.billed ? <Tag tone="success">Billed</Tag> : <Tag tone="outline">Open</Tag>,
                },
                {
                  key: "done",
                  header: "Processed",
                  align: "right",
                  render: (r) => counts(r).processed,
                },
                {
                  key: "left",
                  header: "Unprocessed",
                  align: "right",
                  render: (r) => {
                    const n = counts(r).remaining;
                    return (
                      <span className={cx("font-semibold", n ? "text-grey-900" : "text-cyan-700")}>
                        {n}
                      </span>
                    );
                  },
                },
                {
                  key: "act",
                  header: "",
                  align: "right",
                  render: (r) =>
                    cutoff?.clientId === r.clientId && cutoffKey === r.key ? (
                      <span className="text-xs font-medium text-blue-700">Selected</span>
                    ) : (
                      <Button
                        variant="quiet"
                        className="h-7"
                        onClick={() =>
                          selectCutoff({ clientId: r.clientId, start: r.start, end: r.end })
                        }
                      >
                        Select
                      </Button>
                    ),
                },
              ]}
            />
          </Panel>

          <Panel>
            <PanelHeader
              title="Still to process"
              meta={cutoffClient ? `${cutoffClient.name}, ${cutoffKey}` : "No cutoff selected"}
            />
            {!cutoff ? (
              <EmptyState
                action={<Button onClick={() => navigate("cutoff")}>Select cutoff</Button>}
              >
                Select a cutoff to see who still needs payroll.
              </EmptyState>
            ) : billed ? (
              <EmptyState>
                This cutoff is billed. Payroll is in history and nothing is left to process.
              </EmptyState>
            ) : (
              <DataTable
                rows={queue}
                rowKey={(r) => r.e.id}
                onSelect={payroll ? (r) => navigate(`payroll-entry/${r.e.id}`) : undefined}
                empty={
                  <EmptyState>
                    Everyone in this cutoff is processed. Move to master closes it from Payroll
                    entry.
                  </EmptyState>
                }
                columns={[
                  { key: "no", header: "No.", width: "72px", render: (r) => r.e.empNo },
                  {
                    key: "name",
                    header: "Employee",
                    render: (r) => <span className="font-medium">{fullName(r.e)}</span>,
                  },
                  {
                    key: "pos",
                    header: "Position",
                    render: (r) => <span className="text-grey-600">{r.e.position}</span>,
                  },
                  {
                    key: "time",
                    header: "Time records",
                    render: (r) =>
                      r.finalized ? (
                        <Tag tone="success">Finalized</Tag>
                      ) : r.computed ? (
                        <Tag tone="blue">Computed, not finalized</Tag>
                      ) : (
                        <Tag tone="error">Not computed</Tag>
                      ),
                  },
                  {
                    key: "pay",
                    header: "Payroll",
                    render: () => <Tag tone="outline">Unprocessed</Tag>,
                  },
                  {
                    key: "go",
                    header: "",
                    align: "right",
                    width: "32px",
                    render: () =>
                      payroll ? (
                        <Icon.chevronRight className="ml-auto size-4 text-grey-400 group-hover:text-blue-600" />
                      ) : null,
                  },
                ]}
              />
            )}
          </Panel>
        </div>

        <Panel>
          <PanelHeader title="Shortcuts" />
          {shortcuts.length ? (
            <ul className="divide-y divide-grey-100">
              {shortcuts.map((sc) => (
                <li key={sc.to} className="flex items-center gap-3 px-4 py-3">
                  <div className="min-w-0 flex-1">
                    <div className="text-sm font-medium text-grey-900">{sc.label}</div>
                    <div className="text-xs text-grey-500">{sc.note}</div>
                  </div>
                  <Button
                    variant={sc.primary ? "primary" : "secondary"}
                    onClick={() => navigate(sc.to)}
                  >
                    Open
                  </Button>
                </li>
              ))}
            </ul>
          ) : (
            <EmptyState>
              Your account has no time keeping or payroll rights. Use the menu for the areas you can
              open.
            </EmptyState>
          )}
        </Panel>
      </div>
    </>
  );
}
