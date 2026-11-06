import { useMemo, useState, type ReactNode } from "react";
import { useSession } from "../app/session";
import type { PayrollRecord } from "../data/types";
import { MONTHS, canSee, cutoffHalf, fullName } from "../data/rules";
import { r2 } from "../data/money";
import {
  Button,
  DataTable,
  EmptyState,
  Field,
  Notice,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  TextInput,
  cx,
  type Col,
} from "../ui/kit";
import { Icon } from "../ui/icons";
import { downloadCsv, fmtDate, fmtDateTime, money } from "../ui/format";

type ReportKey = "payroll" | "sss" | "philhealth" | "hdmf" | "hired" | "eoc" | "user-history";
const REPORTS: {
  key: ReportKey;
  label: string;
  group: string;
  right: "reports" | "administrative";
}[] = [
  { key: "payroll", label: "Payroll summary", group: "Payroll", right: "reports" },
  { key: "sss", label: "SSS report", group: "Contributions", right: "reports" },
  { key: "philhealth", label: "PhilHealth report", group: "Contributions", right: "reports" },
  { key: "hdmf", label: "HDMF (Pag-IBIG) report", group: "Contributions", right: "reports" },
  { key: "hired", label: "Hired list", group: "HR reports", right: "reports" },
  { key: "eoc", label: "End of contract list", group: "HR reports", right: "reports" },
  {
    key: "user-history",
    label: "User transaction history",
    group: "Administration",
    right: "administrative",
  },
];

type Row = Record<string, string | number>;
type Spec = {
  cols: { key: string; header: string; money?: boolean }[];
  rows: Row[];
  totals?: string[];
};

export function Reports({ param }: { param?: string }) {
  const { db, user, clients } = useSession();
  const allowed = REPORTS.filter((r) => canSee(user, r.right));
  const [report, setReport] = useState<ReportKey>(
    (allowed.find((r) => r.key === param)?.key ?? allowed[0]?.key) as ReportKey,
  );
  const [clientId, setClientId] = useState(clients[0]?.id ?? "");
  const keys = [...new Set(db.cutoffs.filter((c) => c.clientId === clientId).map((c) => c.key))]
    .sort()
    .reverse();
  const [key, setKey] = useState(keys[0] ?? "");
  const [year, setYear] = useState(2025);
  const [month, setMonth] = useState(11);
  const [userFilter, setUserFilter] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const meta = REPORTS.find((r) => r.key === report)!;
  const client = db.clients.find((c) => c.id === clientId);

  const payrollFor = (pred: (r: PayrollRecord) => boolean) =>
    [...db.history, ...db.openPayroll].filter((r) => r.clientId === clientId && pred(r));
  const emp = (id: string) => db.employees.find((e) => e.id === id)!;
  const line = (r: PayrollRecord, k: string) =>
    [...r.earnings, ...r.deductions].find((l) => l.key === k)?.amount ?? 0;
  const monthKey = (r: PayrollRecord) =>
    r.key.startsWith(`${year}${String(month).padStart(2, "0")}`);

  const spec: Spec = useMemo(() => {
    if (report === "payroll") {
      const rows = payrollFor((r) => r.key === key).map((r) => {
        const e = emp(r.empId);
        const contrib = line(r, "sss") + line(r, "philhealth") + line(r, "pagibig");
        return {
          no: e.empNo,
          name: fullName(e),
          basic: line(r, "basic"),
          other: r2(r.gross - line(r, "basic")),
          gross: r.gross,
          contrib: r2(contrib),
          tax: line(r, "tax"),
          otherDed: r2(r.totalDeductions - contrib - line(r, "tax")),
          net: r.net,
        };
      });
      return {
        cols: [
          { key: "no", header: "No." },
          { key: "name", header: "Employee" },
          { key: "basic", header: "Basic pay", money: true },
          { key: "other", header: "Other earnings", money: true },
          { key: "gross", header: "Gross", money: true },
          { key: "contrib", header: "Contributions", money: true },
          { key: "tax", header: "Tax", money: true },
          { key: "otherDed", header: "Other deductions", money: true },
          { key: "net", header: "Net pay", money: true },
        ],
        rows,
        totals: ["basic", "other", "gross", "contrib", "tax", "otherDed", "net"],
      };
    }
    if (report === "sss" || report === "philhealth" || report === "hdmf") {
      const k = report === "sss" ? "sss" : report === "philhealth" ? "philhealth" : "pagibig";
      const byEmp = new Map<string, number>();
      payrollFor(monthKey).forEach((r) =>
        byEmp.set(r.empId, r2((byEmp.get(r.empId) ?? 0) + line(r, k))),
      );
      const rows = [...byEmp.entries()]
        .filter(([, ee]) => ee > 0)
        .map(([id, ee]): Row => {
          const e = emp(id);
          if (report === "sss") {
            const t = db.sss.find((x) => x.ee === ee);
            return {
              no: e.empNo,
              name: fullName(e),
              gov: e.sssNo || "—",
              ee,
              er: t?.er ?? 0,
              ec: t?.ec ?? 0,
              total: r2(ee + (t?.er ?? 0) + (t?.ec ?? 0)),
            };
          }
          return {
            no: e.empNo,
            name: fullName(e),
            gov: (report === "philhealth" ? e.philhealthNo : e.pagibigNo) || "—",
            ee,
            er: ee,
            total: r2(ee * 2),
          };
        });
      const cols = [
        { key: "no", header: "No." },
        { key: "name", header: "Employee" },
        {
          key: "gov",
          header:
            report === "sss"
              ? "SSS number"
              : report === "philhealth"
                ? "PhilHealth number"
                : "Pag-IBIG number",
        },
        { key: "ee", header: "Employee share", money: true },
        { key: "er", header: "Employer share", money: true },
        ...(report === "sss" ? [{ key: "ec", header: "EC share", money: true }] : []),
        { key: "total", header: "Total", money: true },
      ];
      return {
        cols,
        rows,
        totals: report === "sss" ? ["ee", "er", "ec", "total"] : ["ee", "er", "total"],
      };
    }
    if (report === "hired" || report === "eoc") {
      const field = report === "hired" ? "hired" : "contractEnd";
      const prefix = report === "hired" ? `${year}` : `${year}-${String(month).padStart(2, "0")}`;
      const rows = db.employees
        .filter((e) => e.clientId === clientId && e[field]?.startsWith(prefix))
        .sort((a, b) => a[field].localeCompare(b[field]))
        .map((e) => ({
          no: e.empNo,
          name: fullName(e),
          position: e.position,
          dept: e.department,
          date: fmtDate(e[field]),
          status: e.payrollStatus,
        }));
      return {
        cols: [
          { key: "no", header: "No." },
          { key: "name", header: "Employee" },
          { key: "position", header: "Position" },
          { key: "dept", header: "Department" },
          { key: "date", header: report === "hired" ? "Hired" : "Contract end" },
          { key: "status", header: "Payroll status" },
        ],
        rows,
      };
    }
    const rows = db.logs
      .filter(
        (l) =>
          (!userFilter || l.user === userFilter) &&
          (!from || l.at.slice(0, 10) >= from) &&
          (!to || l.at.slice(0, 10) <= to),
      )
      .map((l) => ({ at: fmtDateTime(l.at), user: l.user, action: l.action, record: l.record }));
    return {
      cols: [
        { key: "at", header: "When" },
        { key: "user", header: "User" },
        { key: "action", header: "Action" },
        { key: "record", header: "Record" },
      ],
      rows,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [report, clientId, key, year, month, userFilter, from, to, db]);

  const cols: Col<Row>[] = spec.cols.map((c) => ({
    key: c.key,
    header: c.header,
    align: c.money ? "right" : "left",
    render: (r) =>
      c.money ? (
        <span className={cx(Number(r[c.key]) === 0 && "text-grey-400")}>
          {money(Number(r[c.key]))}
        </span>
      ) : (
        <span className={c.key === "name" ? "font-medium" : undefined}>{r[c.key]}</span>
      ),
  }));
  const total = (k: string) => r2(spec.rows.reduce((a, r) => a + Number(r[k] ?? 0), 0));
  const subtitle =
    report === "payroll"
      ? `${client?.name ?? ""}, cutoff ${key}${key ? ` (half ${cutoffHalf(`${key.slice(0, 4)}-${key.slice(4, 6)}-${key.endsWith("1") ? "15" : "30"}`)})` : ""} · ${db.cutoffs.find((c) => c.clientId === clientId && c.key === key)?.billed ? "from payroll history" : "open payroll, saved rows only"}`
      : report === "user-history"
        ? "All users"
        : report === "hired"
          ? `${client?.name ?? ""}, ${year}`
          : `${client?.name ?? ""}, ${MONTHS[month - 1]} ${year}`;

  const exportCsv = () =>
    downloadCsv(
      `${report}-${new Date().toISOString().slice(0, 10)}.csv`,
      spec.cols.map((c) => c.header),
      spec.rows.map((r) => spec.cols.map((c) => r[c.key] ?? "")),
    );

  const groups = [...new Set(allowed.map((r) => r.group))];

  return (
    <>
      <PageHeader
        section="Reports"
        title={meta.label}
        description="Filter, check the result table, then print or export."
      />
      <div className="grid items-start gap-4 xl:grid-cols-[220px_minmax(0,1fr)]">
        <Panel className="no-print">
          <nav aria-label="Reports" className="py-2">
            {groups.map((g) => (
              <div key={g} className="mb-2">
                <div className="px-4 pb-1 pt-2 text-xs font-medium text-grey-500">{g}</div>
                {allowed
                  .filter((r) => r.group === g)
                  .map((r) => (
                    <button
                      key={r.key}
                      type="button"
                      onClick={() => setReport(r.key)}
                      aria-current={report === r.key ? "page" : undefined}
                      className={cx(
                        "flex h-8 w-full items-center px-4 text-left text-sm",
                        report === r.key
                          ? "bg-blue-50 font-medium text-blue-800 shadow-[inset_3px_0_0_var(--color-blue-600)]"
                          : "text-grey-700 hover:bg-grey-50",
                      )}
                    >
                      {r.label}
                    </button>
                  ))}
              </div>
            ))}
          </nav>
        </Panel>

        <div className="min-w-0 space-y-4">
          <Panel className="no-print">
            <PanelHeader
              title="Filters"
              actions={
                <>
                  <Button onClick={exportCsv} disabled={!spec.rows.length}>
                    Export
                  </Button>
                  <Button
                    variant="primary"
                    icon={<Icon.printer />}
                    onClick={() => window.print()}
                    disabled={!spec.rows.length}
                  >
                    Print
                  </Button>
                </>
              }
            />
            <div className="flex flex-wrap items-end gap-3 px-4 py-3">
              {report !== "user-history" && (
                <Field label="Client" className="w-72">
                  <Select
                    value={clientId}
                    onChange={(e) => {
                      setClientId(e.target.value);
                      setKey(
                        [
                          ...new Set(
                            db.cutoffs
                              .filter((c) => c.clientId === e.target.value)
                              .map((c) => c.key),
                          ),
                        ]
                          .sort()
                          .reverse()[0] ?? "",
                      );
                    }}
                  >
                    {clients.map((c) => (
                      <option key={c.id} value={c.id}>
                        {c.name}
                      </option>
                    ))}
                  </Select>
                </Field>
              )}
              {report === "payroll" && (
                <Field label="Cutoff" className="w-40">
                  <Select value={key} onChange={(e) => setKey(e.target.value)}>
                    {keys.map((k) => (
                      <option key={k}>{k}</option>
                    ))}
                  </Select>
                </Field>
              )}
              {report !== "payroll" && report !== "user-history" && (
                <Field label="Year" className="w-28">
                  <Select value={year} onChange={(e) => setYear(Number(e.target.value))}>
                    {[2026, 2025, 2024].map((y) => (
                      <option key={y}>{y}</option>
                    ))}
                  </Select>
                </Field>
              )}
              {(report === "sss" ||
                report === "philhealth" ||
                report === "hdmf" ||
                report === "eoc") && (
                <Field label="Month" className="w-40">
                  <Select value={month} onChange={(e) => setMonth(Number(e.target.value))}>
                    {MONTHS.map((m, i) => (
                      <option key={m} value={i + 1}>
                        {m}
                      </option>
                    ))}
                  </Select>
                </Field>
              )}
              {report === "user-history" && (
                <>
                  <Field label="User" className="w-48">
                    <Select value={userFilter} onChange={(e) => setUserFilter(e.target.value)}>
                      <option value="">All users</option>
                      {db.users.map((u) => (
                        <option key={u.id}>{u.username}</option>
                      ))}
                    </Select>
                  </Field>
                  <Field label="From" className="w-44">
                    <TextInput type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
                  </Field>
                  <Field label="To" className="w-44">
                    <TextInput type="date" value={to} onChange={(e) => setTo(e.target.value)} />
                  </Field>
                </>
              )}
            </div>
          </Panel>

          {!db.company && report !== "user-history" && (
            <Notice tone="info" title="Report header is blank" className="no-print">
              Save the company name on the Company screen so printed reports carry a header.
            </Notice>
          )}

          <Panel className="print-flat">
            <div className="flex items-end justify-between gap-4 border-b border-grey-100 px-4 py-3">
              <div>
                {db.company?.name && (
                  <div className="text-xs font-medium text-grey-600">{db.company.name}</div>
                )}
                <h2 className="text-base font-semibold text-grey-900">{meta.label}</h2>
                <p className="text-xs text-grey-500">{subtitle}</p>
              </div>
              <span className="text-xs text-grey-500">
                {spec.rows.length} {spec.rows.length === 1 ? "row" : "rows"}
                {report === "user-history" && " · read-only"}
              </span>
            </div>
            <DataTable
              rows={spec.rows}
              rowKey={(r) => JSON.stringify(r)}
              columns={cols}
              empty={<EmptyState>{emptyText(report)}</EmptyState>}
              foot={
                spec.totals && spec.rows.length ? (
                  <tr>
                    {spec.cols.map((c, i) => (
                      <td
                        key={c.key}
                        className={cx(
                          "h-9 border-t border-grey-200 bg-grey-50 px-3 font-semibold",
                          c.money && "text-right",
                        )}
                      >
                        {i === 0
                          ? "Totals"
                          : spec.totals!.includes(c.key)
                            ? money(total(c.key))
                            : null}
                      </td>
                    ))}
                  </tr>
                ) : undefined
              }
            />
          </Panel>
        </div>
      </div>
    </>
  );
}

function emptyText(r: ReportKey): ReactNode {
  switch (r) {
    case "payroll":
      return "No saved payroll for this client and cutoff.";
    case "sss":
    case "philhealth":
    case "hdmf":
      return "No contributions in saved payroll for this client and month. SSS and PhilHealth come from first halves; Pag-IBIG from second halves.";
    case "hired":
      return "No one on this client was hired in that year.";
    case "eoc":
      return "No contracts on this client end in that month.";
    default:
      return "No transactions match these filters.";
  }
}
