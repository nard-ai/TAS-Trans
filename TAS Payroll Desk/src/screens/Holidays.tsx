import { useState } from "react";
import { useSession } from "../app/session";
import type { Holiday } from "../data/types";
import { leaveLabel } from "../data/rules";
import { RecordEditor } from "../ui/RecordEditor";
import { PageHeader, Select, Tag } from "../ui/kit";
import { fmtDate } from "../ui/format";

export function Holidays() {
  const { db } = useSession();
  const years = [...new Set(db.holidays.map((h) => h.date.slice(0, 4)))].sort().reverse();
  const [year, setYear] = useState(years[0] ?? "2025");
  const rows = db.holidays
    .filter((h) => h.date.startsWith(year))
    .sort((a, b) => a.date.localeCompare(b.date));
  return (
    <>
      <PageHeader
        section="Time keeping"
        title="Holidays"
        description="H is special holiday. L is legal holiday. Payroll pays each at its own rate."
      />
      <RecordEditor<Holiday>
        title="Holiday table"
        noun="holiday"
        table="holidays"
        rows={rows}
        describe={(h) => (h.name ? `${h.name}, ${fmtDate(h.date)}` : "holiday")}
        emptyText={`No holidays for ${year}.`}
        blank={() => ({ date: "", name: "", letter: "H" })}
        toolbar={
          <Select
            aria-label="Year"
            value={year}
            onChange={(e) => setYear(e.target.value)}
            className="w-28"
          >
            {years.map((y) => (
              <option key={y}>{y}</option>
            ))}
          </Select>
        }
        columns={[
          { key: "d", header: "Date", render: (h) => fmtDate(h.date) },
          {
            key: "n",
            header: "Name",
            render: (h) => <span className="font-medium">{h.name}</span>,
          },
          {
            key: "l",
            header: "Letter",
            render: (h) => (
              <Tag
                tone={h.letter === "L" ? "blue" : "neutral"}
              >{`${h.letter} · ${leaveLabel(h.letter)}`}</Tag>
            ),
          },
        ]}
        fields={[
          { key: "date", label: "Date", type: "date", required: true, span: 1 },
          {
            key: "letter",
            label: "Letter",
            type: "select",
            required: true,
            span: 1,
            options: [
              { value: "H", label: "H · Special holiday" },
              { value: "L", label: "L · Legal holiday" },
            ],
          },
          { key: "name", label: "Name", type: "text", required: true },
        ]}
      />
    </>
  );
}
