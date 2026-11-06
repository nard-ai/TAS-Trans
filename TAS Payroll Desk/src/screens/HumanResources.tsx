import { useState } from "react";
import { useSession } from "../app/session";
import type {
  Appraisal,
  Assignment,
  Benefit,
  Employee,
  LeaveBalance,
  Offense,
  Training,
} from "../data/types";
import { fullName, isExcluded, leaveLabel } from "../data/rules";
import { RecordEditor } from "../ui/RecordEditor";
import {
  Amount,
  DataTable,
  EmptyState,
  Field,
  KeyValues,
  PageHeader,
  Panel,
  PanelHeader,
  Select,
  Tabs,
  Tag,
} from "../ui/kit";
import { fmtDate, hrs, money } from "../ui/format";

type TabKey =
  | "details"
  | "benefits"
  | "assignments"
  | "trainings"
  | "offenses"
  | "leaves"
  | "appraisals";

export function HumanResources() {
  const { db, clients } = useSession();
  const ids = new Set(clients.map((c) => c.id));
  const people = db.employees
    .filter((e) => ids.has(e.clientId))
    .sort((a, b) => a.lastName.localeCompare(b.lastName));
  const [empId, setEmpId] = useState(people[0]?.id ?? "");
  const [tab, setTab] = useState<TabKey>("details");
  const emp = db.employees.find((e) => e.id === empId);
  const mine = <T extends { empId: string }>(xs: T[]) => xs.filter((x) => x.empId === empId);

  return (
    <>
      <PageHeader
        section="People"
        title="Human resources"
        description="Benefits, assignments, trainings, offenses, leaves, and appraisals for one person."
      />
      <Panel className="mb-4">
        <div className="flex flex-wrap items-end gap-4 px-4 py-3">
          <Field label="Employee" className="w-[360px] max-w-full">
            <Select value={empId} onChange={(e) => setEmpId(e.target.value)}>
              {people.map((e) => (
                <option key={e.id} value={e.id}>
                  {e.empNo} · {fullName(e)}
                </option>
              ))}
            </Select>
          </Field>
          {emp && (
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1 pb-1.5 text-sm text-grey-600">
              <span>{emp.position}</span>
              <span>{db.clients.find((c) => c.id === emp.clientId)?.name}</span>
              {isExcluded(emp) && <Tag tone="error">Excluded from payroll</Tag>}
            </div>
          )}
        </div>
        <Tabs
          value={tab}
          onChange={setTab}
          tabs={[
            { key: "details", label: "Employee details" },
            { key: "benefits", label: "Benefits", count: mine(db.benefits).length },
            { key: "assignments", label: "Assignments", count: mine(db.assignments).length },
            { key: "trainings", label: "Trainings", count: mine(db.trainings).length },
            { key: "offenses", label: "Offenses", count: mine(db.offenses).length },
            { key: "leaves", label: "Leaves", count: mine(db.leaveBalances).length },
            { key: "appraisals", label: "Appraisals", count: mine(db.appraisals).length },
          ]}
        />
      </Panel>

      {!emp ? (
        <Panel>
          <EmptyState>No employees in your clients.</EmptyState>
        </Panel>
      ) : tab === "details" ? (
        <Details emp={emp} />
      ) : tab === "benefits" ? (
        <RecordEditor<Benefit>
          key={`b${empId}`}
          title="Benefits"
          noun="benefit"
          table="benefits"
          rows={mine(db.benefits)}
          blank={() => ({ empId, benefit: "", amount: 0, start: "", end: "" })}
          describe={(r) => r.benefit || "benefit"}
          emptyText="No benefits recorded for this person."
          columns={[
            { key: "b", header: "Benefit", render: (r) => r.benefit },
            {
              key: "a",
              header: "Amount",
              align: "right",
              render: (r) => <Amount value={r.amount} />,
            },
            { key: "s", header: "Start", render: (r) => fmtDate(r.start) },
            { key: "e", header: "End", render: (r) => fmtDate(r.end) },
          ]}
          fields={[
            { key: "benefit", label: "Benefit", type: "text", required: true },
            { key: "amount", label: "Amount", type: "number", span: 1 },
            { key: "start", label: "Start", type: "date", span: 1 },
            { key: "end", label: "End", type: "date", span: 1 },
          ]}
        />
      ) : tab === "assignments" ? (
        <RecordEditor<Assignment>
          key={`a${empId}`}
          title="Assignments"
          noun="assignment"
          table="assignments"
          rows={mine(db.assignments)}
          blank={() => ({ empId, client: "", position: "", start: "", end: "" })}
          describe={(r) => `${r.position || "assignment"}, ${r.client}`}
          emptyText="No assignments recorded for this person."
          columns={[
            { key: "c", header: "Client", render: (r) => r.client },
            { key: "p", header: "Position", render: (r) => r.position },
            { key: "s", header: "Start", render: (r) => fmtDate(r.start) },
            {
              key: "e",
              header: "End",
              render: (r) =>
                r.end ? fmtDate(r.end) : <span className="text-grey-500">Current</span>,
            },
          ]}
          fields={[
            {
              key: "client",
              label: "Client",
              type: "select",
              required: true,
              options: db.clients.map((c) => ({ value: c.name, label: c.name })),
            },
            { key: "position", label: "Position", type: "text", required: true },
            { key: "start", label: "Start", type: "date", span: 1, required: true },
            { key: "end", label: "End", type: "date", span: 1, hint: "Blank while current." },
          ]}
        />
      ) : tab === "trainings" ? (
        <RecordEditor<Training>
          key={`t${empId}`}
          title="Trainings"
          noun="training"
          table="trainings"
          rows={mine(db.trainings)}
          blank={() => ({ empId, title: "", provider: "", date: "", hours: 0 })}
          describe={(r) => r.title || "training"}
          emptyText="No trainings recorded for this person."
          columns={[
            { key: "t", header: "Title", render: (r) => r.title },
            { key: "p", header: "Provider", render: (r) => r.provider },
            { key: "d", header: "Date", render: (r) => fmtDate(r.date) },
            { key: "h", header: "Hours", align: "right", render: (r) => hrs(r.hours) },
          ]}
          fields={[
            { key: "title", label: "Title", type: "text", required: true },
            { key: "provider", label: "Provider", type: "text" },
            { key: "date", label: "Date", type: "date", span: 1 },
            { key: "hours", label: "Hours", type: "number", span: 1 },
          ]}
        />
      ) : tab === "offenses" ? (
        <RecordEditor<Offense>
          key={`o${empId}`}
          title="Offenses"
          noun="offense"
          table="offenses"
          rows={mine(db.offenses)}
          blank={() => ({ empId, date: "", offense: "", action: "" })}
          describe={(r) => r.offense || "offense"}
          emptyText="No offenses recorded for this person."
          columns={[
            { key: "d", header: "Date", render: (r) => fmtDate(r.date) },
            { key: "o", header: "Offense", render: (r) => r.offense },
            { key: "a", header: "Action taken", render: (r) => r.action },
          ]}
          fields={[
            { key: "date", label: "Date", type: "date", required: true, span: 1 },
            { key: "offense", label: "Offense", type: "text", required: true },
            { key: "action", label: "Action taken", type: "text" },
          ]}
        />
      ) : tab === "leaves" ? (
        <div className="space-y-4">
          <RecordEditor<LeaveBalance>
            key={`l${empId}`}
            title="Leave balances"
            noun="leave balance"
            table="leaveBalances"
            rows={mine(db.leaveBalances)}
            blank={() => ({ empId, year: 2025, type: "Sick leave", entitled: 5, used: 0 })}
            describe={(r) => `${r.type} ${r.year}`}
            emptyText="No leave balances for this person."
            maxHeight="320px"
            columns={[
              { key: "y", header: "Year", render: (r) => r.year },
              { key: "t", header: "Type", render: (r) => r.type },
              {
                key: "e",
                header: "Entitled days",
                align: "right",
                render: (r) => r.entitled.toFixed(1),
              },
              { key: "u", header: "Used days", align: "right", render: (r) => r.used.toFixed(1) },
              {
                key: "b",
                header: "Balance",
                align: "right",
                render: (r) => (
                  <span className="font-semibold">{(r.entitled - r.used).toFixed(1)}</span>
                ),
              },
            ]}
            fields={[
              { key: "year", label: "Year", type: "number", decimals: 0, span: 1, required: true },
              {
                key: "type",
                label: "Type",
                type: "select",
                span: 1,
                required: true,
                options: [
                  { value: "Sick leave", label: "Sick leave" },
                  { value: "Vacation leave", label: "Vacation leave" },
                ],
              },
              { key: "entitled", label: "Entitled days", type: "number", span: 1, decimals: 1 },
              { key: "used", label: "Used days", type: "number", span: 1, decimals: 1 },
            ]}
          />
          <Panel>
            <PanelHeader title="Filed leave" meta="File or change leave on Leave and overtime." />
            <DataTable
              rows={mine(db.leaves).sort((a, b) => b.date.localeCompare(a.date))}
              rowKey={(r) => r.id}
              empty={<EmptyState>No leave filed for this person.</EmptyState>}
              columns={[
                { key: "d", header: "Date", render: (r) => fmtDate(r.date) },
                {
                  key: "l",
                  header: "Letter",
                  width: "64px",
                  render: (r) => <span className="font-semibold">{r.letter}</span>,
                },
                { key: "n", header: "Meaning", render: (r) => leaveLabel(r.letter) },
                { key: "h", header: "Hours", align: "right", render: (r) => hrs(r.hours) },
              ]}
            />
          </Panel>
        </div>
      ) : (
        <RecordEditor<Appraisal>
          key={`ap${empId}`}
          title="Appraisals"
          noun="appraisal"
          table="appraisals"
          rows={mine(db.appraisals)}
          blank={() => ({ empId, date: "", rating: "", rater: "", remarks: "" })}
          describe={(r) => `${fmtDate(r.date)} appraisal`}
          emptyText="No appraisals recorded for this person."
          columns={[
            { key: "d", header: "Date", render: (r) => fmtDate(r.date) },
            { key: "r", header: "Rating", render: (r) => r.rating },
            { key: "b", header: "Rater", render: (r) => r.rater },
          ]}
          fields={[
            { key: "date", label: "Date", type: "date", required: true, span: 1 },
            {
              key: "rating",
              label: "Rating",
              type: "select",
              span: 1,
              required: true,
              options: [
                "Outstanding",
                "Very satisfactory",
                "Satisfactory",
                "Needs improvement",
                "Poor",
              ].map((v) => ({ value: v, label: v })),
            },
            { key: "rater", label: "Rater", type: "text" },
            { key: "remarks", label: "Remarks", type: "text" },
          ]}
        />
      )}
    </>
  );
}

function Details({ emp }: { emp: Employee }) {
  const { db } = useSession();
  const client = db.clients.find((c) => c.id === emp.clientId);
  return (
    <Panel>
      <PanelHeader
        title={`${emp.empNo} ${fullName(emp)}`}
        meta="Read only. Edit the profile in Employees."
      />
      <div className="space-y-5 px-4 py-4">
        <KeyValues
          cols={4}
          items={[
            { label: "Employee number", value: emp.empNo },
            { label: "DTR id", value: emp.dtrId },
            { label: "Client", value: client?.name },
            { label: "Department", value: emp.department },
            { label: "Position", value: emp.position },
            { label: "Shift", value: emp.shift },
            {
              label: "Payroll status",
              value: `${emp.payrollStatus}${isExcluded(emp) ? " (excluded)" : ""}`,
            },
            { label: "Active", value: emp.active ? "Yes" : "No" },
          ]}
        />
        <div className="border-t border-grey-100 pt-4">
          <KeyValues
            cols={4}
            items={[
              { label: "Gender", value: emp.gender },
              { label: "Civil status", value: emp.civilStatus },
              { label: "Birth date", value: fmtDate(emp.birthDate) },
              { label: "Address", value: emp.address },
              { label: "Tax code", value: emp.taxCode },
              { label: "TIN", value: emp.tin },
              { label: "SSS number", value: emp.sssNo },
              { label: "PhilHealth number", value: emp.philhealthNo },
            ]}
          />
        </div>
        <div className="border-t border-grey-100 pt-4">
          <KeyValues
            cols={4}
            items={[
              { label: "Payment basis", value: emp.basis },
              { label: "Monthly rate", value: money(emp.monthlyRate) },
              { label: "Daily rate", value: money(emp.dailyRate) },
              {
                label: "Hourly rate",
                value: emp.hourlyRate ? money(emp.hourlyRate) : "Daily rate / 8",
              },
              { label: "Hired", value: fmtDate(emp.hired) },
              { label: "Contract start", value: fmtDate(emp.contractStart) },
              { label: "Contract end", value: fmtDate(emp.contractEnd) },
              { label: "Bank", value: `${emp.bank} ${emp.accountNo}`.trim() },
            ]}
          />
        </div>
      </div>
    </Panel>
  );
}
