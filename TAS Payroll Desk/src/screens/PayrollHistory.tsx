import { useMemo, useState } from "react";
import { useSession } from "../app/session";
import { fullName } from "../data/rules";
import {
  Amount,
  Button,
  DataTable,
  EmptyState,
  Field,
  PageHeader,
  Panel,
  PanelHeader,
  SearchInput,
  Select,
  Tag,
} from "../ui/kit";
import { Icon } from "../ui/icons";
import { fmtRange } from "../ui/format";
import { PayslipDoc } from "./PayslipDoc";

export function PayrollHistory() {
  const { db, clients } = useSession();
  const ids = new Set(clients.map((c) => c.id));
  const rows = db.history.filter((r) => ids.has(r.clientId));
  const [clientId, setClientId] = useState(
    clients.find((c) => rows.some((r) => r.clientId === c.id))?.id ?? "",
  );
  const keys = [...new Set(rows.filter((r) => r.clientId === clientId).map((r) => r.key))]
    .sort()
    .reverse();
  const [key, setKey] = useState(keys[0] ?? "");
  const [q, setQ] = useState("");
  const [sel, setSel] = useState<string | null>(null);

  const list = useMemo(() => {
    const term = q.trim().toLowerCase();
    return rows
      .filter((r) => r.clientId === clientId && (!key || r.key === key))
      .map((r) => ({ r, e: db.employees.find((e) => e.id === r.empId)! }))
      .filter(
        ({ e }) =>
          e && (!term || e.empNo.includes(term) || fullName(e).toLowerCase().includes(term)),
      )
      .sort((a, b) => a.e.lastName.localeCompare(b.e.lastName));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [db.history, clientId, key, q]);
  const open = list.find((x) => x.r.id === sel);
  const client = db.clients.find((c) => c.id === clientId);
  const gross = list.reduce((a, x) => a + x.r.gross, 0);
  const net = list.reduce((a, x) => a + x.r.net, 0);

  return (
    <>
      <PageHeader
        section="Payroll"
        title="Payroll history"
        description="Closed payroll after Move to master. Read-only. Payslips print from here."
      />
      <Panel className="no-print mb-4">
        <div className="grid gap-3 px-4 py-3 sm:grid-cols-[1fr_200px_1fr]">
          <Field label="Client">
            <Select
              value={clientId}
              onChange={(e) => {
                setClientId(e.target.value);
                const k = [
                  ...new Set(rows.filter((r) => r.clientId === e.target.value).map((r) => r.key)),
                ]
                  .sort()
                  .reverse();
                setKey(k[0] ?? "");
                setSel(null);
              }}
            >
              {clients.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.name}
                </option>
              ))}
            </Select>
          </Field>
          <Field label="Cutoff">
            <Select value={key} onChange={(e) => (setKey(e.target.value), setSel(null))}>
              {keys.length === 0 && <option value="">None</option>}
              {keys.map((k) => (
                <option key={k}>{k}</option>
              ))}
            </Select>
          </Field>
          <Field label="Employee">
            <SearchInput
              label="Search employee"
              value={q}
              onChange={setQ}
              placeholder="Name or number"
            />
          </Field>
        </div>
      </Panel>

      <div className="grid items-start gap-4 2xl:grid-cols-[minmax(0,1fr)_minmax(0,720px)]">
        <Panel className="no-print">
          <PanelHeader title="Closed payroll" meta={`${list.length} rows`} />
          <DataTable
            rows={list}
            rowKey={(x) => x.r.id}
            selected={sel}
            onSelect={(x) => setSel(x.r.id)}
            maxHeight="calc(100vh - 340px)"
            empty={
              <EmptyState>
                No closed payroll for this client yet. Payroll appears here after Move to master.
              </EmptyState>
            }
            columns={[
              {
                key: "c",
                header: "Client",
                render: () => <span className="text-grey-600">{client?.code}</span>,
              },
              {
                key: "k",
                header: "Cutoff",
                render: (x) => <span title={fmtRange(x.r.start, x.r.end)}>{x.r.key}</span>,
              },
              {
                key: "e",
                header: "Employee",
                render: (x) => (
                  <span className="font-medium">{`${x.e.empNo} ${fullName(x.e)}`}</span>
                ),
              },
              {
                key: "g",
                header: "Gross",
                align: "right",
                render: (x) => <Amount value={x.r.gross} />,
              },
              {
                key: "n",
                header: "Net",
                align: "right",
                render: (x) => <Amount value={x.r.net} strong />,
              },
              {
                key: "p",
                header: "Processed by",
                render: (x) => <span className="text-grey-600">{x.r.processedBy}</span>,
              },
              {
                key: "l",
                header: "",
                render: () => (
                  <Tag tone="success">
                    <Icon.lock className="mr-1 size-3" />
                    Closed
                  </Tag>
                ),
              },
            ]}
            foot={
              list.length > 0 ? (
                <tr>
                  <td
                    colSpan={3}
                    className="h-9 border-t border-grey-200 bg-grey-50 px-3 text-sm font-semibold"
                  >
                    Totals
                  </td>
                  <td className="border-t border-grey-200 bg-grey-50 px-3 text-right font-semibold">
                    <Amount value={gross} strong />
                  </td>
                  <td className="border-t border-grey-200 bg-grey-50 px-3 text-right">
                    <Amount value={net} strong />
                  </td>
                  <td colSpan={2} className="border-t border-grey-200 bg-grey-50" />
                </tr>
              ) : undefined
            }
          />
        </Panel>

        <div className="min-w-0">
          {open && client ? (
            <>
              <div className="no-print mb-2 flex items-center justify-between">
                <span className="text-sm text-grey-600">Read-only. No Save.</span>
                <Button variant="primary" icon={<Icon.printer />} onClick={() => window.print()}>
                  Print payslip
                </Button>
              </div>
              <PayslipDoc
                record={open.r}
                emp={open.e}
                client={client}
                company={db.company}
                source="history"
              />
            </>
          ) : (
            <Panel className="no-print">
              <EmptyState>Choose a row to open the closed payslip.</EmptyState>
            </Panel>
          )}
        </div>
      </div>
    </>
  );
}
