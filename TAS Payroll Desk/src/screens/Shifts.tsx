import { useSession } from "../app/session";
import type { Shift } from "../data/types";
import { RecordEditor } from "../ui/RecordEditor";
import { PageHeader, Tag } from "../ui/kit";

export function Shifts() {
  const { db } = useSession();
  return (
    <>
      <PageHeader
        section="Administration"
        title="Shift schedule"
        description="One row per shift. These are the fields the time-record batch reads."
      />
      <RecordEditor<Shift>
        title="Shifts"
        noun="shift"
        table="shifts"
        rows={db.shifts}
        describe={(s) => s.code || "shift"}
        emptyText="No shifts yet. The time-record batch needs at least one."
        blank={() => ({
          code: "",
          start: "08:00",
          end: "17:00",
          midday: "12:00",
          workHours: 8,
          graceMinutes: 10,
          graceOn: true,
          noLate: false,
          noUndertime: false,
          undertimeMinutes: 0,
          noInOut: false,
          flexible: false,
          noOvertime: false,
          overtimeStart: "17:30",
          nightDiff: false,
        })}
        columns={[
          {
            key: "code",
            header: "Code",
            render: (s) => <span className="font-medium">{s.code}</span>,
          },
          { key: "time", header: "Start – end", render: (s) => `${s.start} – ${s.end}` },
          {
            key: "wh",
            header: "Work hours",
            align: "right",
            render: (s) => s.workHours.toFixed(2),
          },
          {
            key: "grace",
            header: "Grace",
            align: "right",
            render: (s) =>
              s.graceOn ? `${s.graceMinutes} min` : <span className="text-grey-400">Off</span>,
          },
          {
            key: "ot",
            header: "Overtime from",
            render: (s) =>
              s.noOvertime ? <span className="text-grey-400">No overtime</span> : s.overtimeStart,
          },
          {
            key: "nd",
            header: "Night diff",
            render: (s) =>
              s.nightDiff ? <Tag tone="blue">On</Tag> : <span className="text-grey-400">Off</span>,
          },
        ]}
        fields={[
          { key: "code", label: "Code", type: "text", required: true, span: 1 },
          { key: "workHours", label: "Work hours", type: "number", span: 1 },
          { key: "start", label: "Start", type: "time", span: 1, required: true },
          { key: "end", label: "End", type: "time", span: 1, required: true },
          {
            key: "midday",
            label: "Midday",
            type: "time",
            span: 1,
            hint: "One hour is taken for a meal when work spans this time.",
          },
          { key: "overtimeStart", label: "Overtime start", type: "time", span: 1 },
          { key: "graceMinutes", label: "Grace minutes", type: "number", decimals: 0, span: 1 },
          {
            key: "undertimeMinutes",
            label: "Undertime minutes",
            type: "number",
            decimals: 0,
            span: 1,
            hint: "Undertime at or under this is ignored.",
          },
          { key: "graceOn", label: "Grace on", type: "checkbox" },
          { key: "noLate", label: "No late", type: "checkbox" },
          { key: "noUndertime", label: "No undertime", type: "checkbox" },
          {
            key: "noInOut",
            label: "No in/out",
            type: "checkbox",
            hint: "Treat working days as present without punches.",
          },
          { key: "flexible", label: "Flexible time", type: "checkbox" },
          { key: "noOvertime", label: "No overtime", type: "checkbox" },
          { key: "nightDiff", label: "Night differential", type: "checkbox" },
        ]}
      />
    </>
  );
}
