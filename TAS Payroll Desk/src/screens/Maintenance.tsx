import { useState } from "react";
import { useSession } from "../app/session";
import type {
  Bank,
  Branch,
  Department,
  HdmfRow,
  PhRow,
  SssRow,
  TaxCode,
  TaxRow,
} from "../data/types";
import { RecordEditor } from "../ui/RecordEditor";
import { Amount, PageHeader, Panel, Select, Tabs } from "../ui/kit";

type K = "banks" | "branches" | "departments" | "sss" | "philhealth" | "hdmf" | "tax" | "taxCodes";

export function Maintenance() {
  const { db } = useSession();
  const [tab, setTab] = useState<K>("banks");
  const [taxTable, setTaxTable] = useState<"Semi-monthly" | "2018">("2018");
  const [taxCode, setTaxCode] = useState("");
  const m = (n: number) => <Amount value={n} />;
  const range = (from: number, to: number) => (
    <span>
      {from.toLocaleString("en-PH", { minimumFractionDigits: 2 })} –{" "}
      {to >= 9999999 ? "and over" : to.toLocaleString("en-PH", { minimumFractionDigits: 2 })}
    </span>
  );
  const taxRows = db.tax
    .filter(
      (t) => t.table === taxTable && (taxTable === "2018" || !taxCode || t.taxCode === taxCode),
    )
    .sort((a, b) => a.taxCode.localeCompare(b.taxCode) || a.floor - b.floor);
  const codes = [
    ...new Set(db.tax.filter((t) => t.table === "Semi-monthly").map((t) => t.taxCode)),
  ];

  return (
    <>
      <PageHeader
        section="Administration"
        title="Maintenance tables"
        description="Rows the server uses to compute pay. Edit rows here; nothing on this screen previews a payroll."
      />
      <Panel className="mb-4">
        <Tabs
          value={tab}
          onChange={setTab}
          tabs={[
            { key: "banks", label: "Banks", count: db.banks.length },
            { key: "branches", label: "Branches", count: db.branches.length },
            { key: "departments", label: "Departments", count: db.departments.length },
            { key: "sss", label: "SSS table", count: db.sss.length },
            { key: "philhealth", label: "PhilHealth table", count: db.philhealth.length },
            { key: "hdmf", label: "Pag-IBIG table", count: db.hdmf.length },
            { key: "tax", label: "Tax table", count: db.tax.length },
            { key: "taxCodes", label: "Tax exemption codes", count: db.taxCodes.length },
          ]}
        />
      </Panel>
      {tab === "banks" && (
        <RecordEditor<Bank>
          key="banks"
          title="Banks"
          noun="bank"
          table="banks"
          rows={db.banks}
          describe={(r) => r.name || "bank"}
          emptyText="No banks yet."
          blank={() => ({ code: "", name: "" })}
          columns={[
            {
              key: "c",
              header: "Code",
              render: (r) => <span className="font-medium">{r.code}</span>,
            },
            { key: "n", header: "Name", render: (r) => r.name },
          ]}
          fields={[
            { key: "code", label: "Code", type: "text", required: true, span: 1 },
            { key: "name", label: "Name", type: "text", required: true },
          ]}
        />
      )}
      {tab === "branches" && (
        <RecordEditor<Branch>
          key="branches"
          title="Branches"
          noun="branch"
          table="branches"
          rows={db.branches}
          describe={(r) => r.name || "branch"}
          emptyText="No branches yet."
          blank={() => ({ code: "", name: "", bank: "" })}
          columns={[
            {
              key: "c",
              header: "Code",
              render: (r) => <span className="font-medium">{r.code}</span>,
            },
            { key: "n", header: "Name", render: (r) => r.name },
            { key: "b", header: "Bank", render: (r) => r.bank },
          ]}
          fields={[
            { key: "code", label: "Code", type: "text", required: true, span: 1 },
            {
              key: "bank",
              label: "Bank",
              type: "select",
              span: 1,
              options: db.banks.map((b) => ({ value: b.code, label: b.name })),
            },
            { key: "name", label: "Name", type: "text", required: true },
          ]}
        />
      )}
      {tab === "departments" && (
        <RecordEditor<Department>
          key="departments"
          title="Departments"
          noun="department"
          table="departments"
          rows={db.departments}
          describe={(r) => r.name || "department"}
          emptyText="No departments yet."
          blank={() => ({ code: "", name: "" })}
          columns={[
            {
              key: "c",
              header: "Code",
              render: (r) => <span className="font-medium">{r.code}</span>,
            },
            { key: "n", header: "Name", render: (r) => r.name },
          ]}
          fields={[
            { key: "code", label: "Code", type: "text", required: true, span: 1 },
            { key: "name", label: "Name", type: "text", required: true },
          ]}
        />
      )}
      {tab === "sss" && (
        <RecordEditor<SssRow>
          key="sss"
          title="SSS table"
          noun="SSS row"
          table="sss"
          rows={[...db.sss].sort((a, b) => a.from - b.from)}
          describe={(r) => `wage from ${r.from.toLocaleString("en-PH")}`}
          emptyText="No SSS rows. SSS deductions will be zero."
          blank={() => ({ from: 0, to: 0, ee: 0, er: 0, ec: 0 })}
          columns={[
            { key: "w", header: "Wage range", render: (r) => range(r.from, r.to) },
            { key: "ee", header: "Employee", align: "right", render: (r) => m(r.ee) },
            { key: "er", header: "Employer", align: "right", render: (r) => m(r.er) },
            { key: "ec", header: "EC", align: "right", render: (r) => m(r.ec) },
          ]}
          fields={[
            { key: "from", label: "Wage from", type: "number", required: true, span: 1 },
            { key: "to", label: "Wage to", type: "number", required: true, span: 1 },
            { key: "ee", label: "Employee share", type: "number", span: 1 },
            { key: "er", label: "Employer share", type: "number", span: 1 },
            { key: "ec", label: "EC share", type: "number", span: 1 },
          ]}
        />
      )}
      {tab === "philhealth" && (
        <RecordEditor<PhRow>
          key="ph"
          title="PhilHealth table"
          noun="PhilHealth row"
          table="philhealth"
          rows={[...db.philhealth].sort((a, b) => a.from - b.from)}
          describe={(r) => `wage from ${r.from.toLocaleString("en-PH")}`}
          emptyText="No PhilHealth rows. PhilHealth deductions will be zero."
          blank={() => ({ from: 0, to: 0, ee: 0, er: 0 })}
          columns={[
            { key: "w", header: "Wage range", render: (r) => range(r.from, r.to) },
            { key: "ee", header: "Employee", align: "right", render: (r) => m(r.ee) },
            { key: "er", header: "Employer", align: "right", render: (r) => m(r.er) },
          ]}
          fields={[
            { key: "from", label: "Wage from", type: "number", required: true, span: 1 },
            { key: "to", label: "Wage to", type: "number", required: true, span: 1 },
            { key: "ee", label: "Employee share", type: "number", span: 1 },
            { key: "er", label: "Employer share", type: "number", span: 1 },
          ]}
        />
      )}
      {tab === "hdmf" && (
        <RecordEditor<HdmfRow>
          key="hdmf"
          title="Pag-IBIG table"
          noun="Pag-IBIG row"
          table="hdmf"
          rows={[...db.hdmf].sort((a, b) => a.from - b.from)}
          describe={(r) => `wage from ${r.from.toLocaleString("en-PH")}`}
          emptyText="No Pag-IBIG rows. Pag-IBIG deductions will be zero."
          blank={() => ({ from: 0, to: 0, eeRate: 2, erRate: 2 })}
          columns={[
            { key: "w", header: "Wage range", render: (r) => range(r.from, r.to) },
            {
              key: "ee",
              header: "Employee rate",
              align: "right",
              render: (r) => `${r.eeRate.toFixed(2)}%`,
            },
            {
              key: "er",
              header: "Employer rate",
              align: "right",
              render: (r) => `${r.erRate.toFixed(2)}%`,
            },
          ]}
          fields={[
            { key: "from", label: "Wage from", type: "number", required: true, span: 1 },
            { key: "to", label: "Wage to", type: "number", required: true, span: 1 },
            { key: "eeRate", label: "Employee rate (percent)", type: "number", span: 1 },
            { key: "erRate", label: "Employer rate (percent)", type: "number", span: 1 },
          ]}
        />
      )}
      {tab === "tax" && (
        <RecordEditor<TaxRow>
          key={`tax-${taxTable}`}
          title="Tax table"
          noun="tax row"
          table="tax"
          rows={taxRows}
          describe={(r) => `${r.table} ${r.taxCode}, floor ${r.floor.toLocaleString("en-PH")}`}
          emptyText="No rows for this table and code. Payroll will stop with a missing tax code error."
          blank={() => ({
            table: taxTable,
            taxCode: taxTable === "2018" ? "All codes" : taxCode,
            floor: 0,
            percent: 0,
            baseTax: 0,
          })}
          toolbar={
            <>
              <Select
                aria-label="Tax table"
                value={taxTable}
                onChange={(e) => setTaxTable(e.target.value as "Semi-monthly" | "2018")}
                className="w-40"
              >
                <option value="2018">2018 rows</option>
                <option value="Semi-monthly">Semi-monthly rows</option>
              </Select>
              {taxTable === "Semi-monthly" && (
                <Select
                  aria-label="Tax code"
                  value={taxCode}
                  onChange={(e) => setTaxCode(e.target.value)}
                  className="w-36"
                >
                  <option value="">All codes</option>
                  {codes.map((c) => (
                    <option key={c}>{c}</option>
                  ))}
                </Select>
              )}
            </>
          }
          columns={[
            {
              key: "c",
              header: "Tax code",
              render: (r) => <span className="font-medium">{r.taxCode}</span>,
            },
            { key: "f", header: "Bracket floor", align: "right", render: (r) => m(r.floor) },
            { key: "p", header: "Percent", align: "right", render: (r) => `${r.percent}%` },
            { key: "b", header: "Base tax", align: "right", render: (r) => m(r.baseTax) },
          ]}
          fields={[
            {
              key: "table",
              label: "Table",
              type: "select",
              span: 1,
              required: true,
              options: [
                { value: "Semi-monthly", label: "Semi-monthly" },
                { value: "2018", label: "2018" },
              ],
            },
            { key: "taxCode", label: "Tax code", type: "text", span: 1, required: true },
            { key: "floor", label: "Bracket floor", type: "number", span: 1 },
            { key: "percent", label: "Percent", type: "number", span: 1 },
            { key: "baseTax", label: "Base tax", type: "number", span: 1 },
          ]}
        />
      )}
      {tab === "taxCodes" && (
        <RecordEditor<TaxCode>
          key="taxCodes"
          title="Tax exemption codes"
          noun="tax code"
          table="taxCodes"
          rows={db.taxCodes}
          describe={(r) => r.code || "tax code"}
          emptyText="No tax codes yet."
          blank={() => ({ code: "", description: "", exemption: 0 })}
          columns={[
            {
              key: "c",
              header: "Code",
              render: (r) => <span className="font-medium">{r.code}</span>,
            },
            { key: "d", header: "Description", render: (r) => r.description },
            { key: "e", header: "Exemption", align: "right", render: (r) => m(r.exemption) },
            {
              key: "t",
              header: "In tax table",
              render: (r) =>
                db.tax.some((t) => t.table === "Semi-monthly" && t.taxCode === r.code) ? (
                  "Yes"
                ) : (
                  <span className="font-medium text-red-600">Missing</span>
                ),
            },
          ]}
          fields={[
            { key: "code", label: "Code", type: "text", required: true, span: 1 },
            { key: "exemption", label: "Exemption", type: "number", span: 1 },
            { key: "description", label: "Description", type: "text" },
          ]}
        />
      )}
    </>
  );
}
