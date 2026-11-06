import type { ReactNode } from "react";
import { useSession } from "../app/session";
import { navigate } from "../app/router";
import type { Employee, StaffRemark } from "../data/types";
import { fullName, isExcluded, isFmsClient } from "../data/rules";
import { Button, DataTable, EmptyState, Notice, Panel, PanelHeader, Tag, cx } from "../ui/kit";

export function RemarkCell({ remark }: { remark: StaffRemark }) {
  if (remark === "Absent") return <span className="font-medium text-red-600">Absent</span>;
  if (remark === "No In/Out") return <Tag tone="error">No In/Out</Tag>;
  if (remark === "Invalid Day") return <span className="text-grey-400">Invalid Day</span>;
  if (remark === "Day off") return <span className="text-grey-400">Day off</span>;
  return <span className="text-grey-700">Present</span>;
}

/** Guards shared by both time keeping screens. Returns a blocking view or null. */
export function useTimeGuard(kind: "staff" | "fms"): ReactNode | null {
  const { cutoff, cutoffClient } = useSession();
  if (!cutoff || !cutoffClient)
    return (
      <Notice
        tone="error"
        title="No cutoff selected"
        actions={<Button onClick={() => navigate("cutoff")}>Select cutoff</Button>}
      >
        Time keeping needs a client and period. The grid stays closed until a cutoff is selected.
      </Notice>
    );
  const fms = isFmsClient(cutoffClient);
  if (kind === "staff" && fms)
    return (
      <Notice
        tone="info"
        title={`${cutoffClient.name} is an FMS client`}
        actions={<Button onClick={() => navigate("time-fms")}>Open driver time keeping</Button>}
      >
        Staff time keeping is shown only for a staff client. Drivers and conductors use their own
        grid.
      </Notice>
    );
  if (kind === "fms" && !fms)
    return (
      <Notice
        tone="info"
        title={`${cutoffClient.name} is a staff client`}
        actions={<Button onClick={() => navigate("time-staff")}>Open staff time keeping</Button>}
      >
        Driver time keeping is shown only for an FMS client.
      </Notice>
    );
  return null;
}

export function TimeEmployeeList({
  kind,
  selected,
  onSelect,
}: {
  kind: "staff" | "fms";
  selected: string | null;
  onSelect: (e: Employee) => void;
}) {
  const { db, cutoffClient, cutoffKey } = useSession();
  const final = kind === "staff" ? db.staffFinal : db.fmsFinal;
  const people = db.employees
    .filter((e) => e.clientId === cutoffClient?.id && !isExcluded(e))
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const stateOf = (e: Employee) => {
    const rows = final[`${e.id}|${cutoffKey}`];
    const fin = db.payrollAttendance[`${e.id}|${cutoffKey}`];
    if (!rows) return "none";
    if (!fin) return "computed";
    const finRows = kind === "staff" ? fin.staff : fin.fms;
    const same =
      kind === "staff"
        ? JSON.stringify(finRows) ===
          JSON.stringify(rows.filter((r) => "remark" in r && r.remark !== "Invalid Day"))
        : JSON.stringify(finRows) === JSON.stringify(rows);
    return same ? "final" : "changed";
  };
  const done = people.filter((e) => stateOf(e) === "final").length;
  return (
    <Panel>
      <PanelHeader
        title={kind === "staff" ? "Employees" : "Drivers and conductors"}
        meta={`${done} of ${people.length} finalized`}
      />
      <DataTable
        rows={people}
        rowKey={(e) => e.id}
        selected={selected}
        onSelect={onSelect}
        maxHeight="calc(100vh - 260px)"
        empty={<EmptyState>No employees on this client are included in payroll.</EmptyState>}
        columns={[
          {
            key: "n",
            header: "Employee and status",
            render: (e) => {
              const s = stateOf(e);
              return (
                <div className="leading-tight">
                  <div className="truncate font-medium text-grey-900">{fullName(e)}</div>
                  <div className="mt-0.5 flex items-center justify-between gap-2 text-xs">
                    <span className="text-grey-500">
                      {e.empNo}
                      {e.role ? ` · ${e.role}` : ""}
                    </span>
                    <span
                      className={cx(
                        "font-medium",
                        s === "final"
                          ? "text-cyan-800"
                          : s === "none"
                            ? "text-grey-400"
                            : "text-blue-700",
                      )}
                    >
                      {s === "final"
                        ? "Finalized"
                        : s === "changed"
                          ? "Changed"
                          : s === "computed"
                            ? "Computed"
                            : "Not computed"}
                    </span>
                  </div>
                </div>
              );
            },
          },
        ]}
      />
    </Panel>
  );
}

export function TotalsRow({ cells }: { cells: (ReactNode | null)[] }) {
  return (
    <tr>
      {cells.map((c, i) => (
        <td
          key={i}
          className={cx(
            "h-9 whitespace-nowrap border-t border-grey-200 bg-grey-50 px-2.5 text-sm font-semibold text-grey-900",
            i > 0 && "text-right",
          )}
        >
          {c}
        </td>
      ))}
    </tr>
  );
}
