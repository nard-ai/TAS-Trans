import { useState } from "react";
import { useSession } from "../app/session";
import { navigate, previousRoute } from "../app/router";
import { fullName, isExcluded } from "../data/rules";
import { Button, EmptyState, Field, Notice, PageHeader, Panel, Select } from "../ui/kit";
import { Icon } from "../ui/icons";
import { PayslipDoc } from "./PayslipDoc";
import { fmtRange } from "../ui/format";

export function PayslipScreen({ param }: { param?: string }) {
  const { db, cutoff, cutoffClient, cutoffKey, billed } = useSession();
  const people = cutoffClient
    ? db.employees
        .filter((e) => e.clientId === cutoffClient.id && !isExcluded(e))
        .sort((a, b) => a.lastName.localeCompare(b.lastName))
    : [];
  const [empId, setEmpId] = useState(param ?? people[0]?.id ?? "");
  const emp = db.employees.find((e) => e.id === empId);
  const source = billed ? "history" : "open";
  const record = (billed ? db.history : db.openPayroll).find(
    (r) => r.empId === empId && r.key === cutoffKey && r.clientId === cutoffClient?.id,
  );

  return (
    <>
      <PageHeader
        section="Payroll"
        title="Payslip"
        description="Prints what was saved. It does not recompute."
        actions={
          <>
            <Button
              onClick={() =>
                navigate(previousRoute() === "payslip" ? "payroll-entry" : previousRoute())
              }
            >
              Close
            </Button>
            <Button
              variant="primary"
              icon={<Icon.printer />}
              onClick={() => window.print()}
              disabled={!record}
            >
              Print
            </Button>
          </>
        }
      />
      {!cutoff || !cutoffClient ? (
        <Notice
          tone="error"
          title="No cutoff selected"
          actions={<Button onClick={() => navigate("cutoff")}>Select cutoff</Button>}
        >
          Payslips print for the selected cutoff. Use Payroll history for closed cutoffs.
        </Notice>
      ) : (
        <>
          <Panel className="no-print mb-4">
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
              <p className="pb-1.5 text-sm text-grey-600">
                {cutoffClient.name}, {fmtRange(cutoff.start, cutoff.end)}
              </p>
            </div>
          </Panel>
          {record && emp ? (
            <PayslipDoc
              record={record}
              emp={emp}
              client={cutoffClient}
              company={db.company}
              source={source}
            />
          ) : (
            <Panel>
              <EmptyState
                action={
                  <Button onClick={() => navigate(`payroll-entry/${empId}`)}>
                    Open payroll entry
                  </Button>
                }
              >
                No saved payroll for this employee and cutoff.
              </EmptyState>
            </Panel>
          )}
        </>
      )}
    </>
  );
}
