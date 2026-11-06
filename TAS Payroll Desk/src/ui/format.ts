import { parseISO } from "../data/rules";

const M = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export const money = (n: number | null | undefined) =>
  n == null
    ? "—"
    : n.toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
export const hrs = (n: number | null | undefined) => (n == null ? "—" : n.toFixed(2));

export function fmtDate(iso: string | null | undefined) {
  if (!iso) return "—";
  const { y, m, d } = parseISO(iso.slice(0, 10));
  return `${M[m - 1]} ${d}, ${y}`;
}
export function fmtShort(iso: string) {
  const { m, d } = parseISO(iso.slice(0, 10));
  return `${M[m - 1]} ${d}`;
}
export function fmtRange(start: string, end: string) {
  const a = parseISO(start);
  const b = parseISO(end);
  if (a.y === b.y && a.m === b.m) return `${M[a.m - 1]} ${a.d}–${b.d}, ${a.y}`;
  return `${fmtDate(start)} – ${fmtDate(end)}`;
}
export function fmtDateTime(iso: string | null | undefined) {
  if (!iso) return "—";
  const time = iso.slice(11, 16);
  return `${fmtDate(iso)} ${time}`;
}

export function downloadCsv(name: string, header: string[], rows: (string | number)[][]) {
  const esc = (v: string | number) => {
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const csv = [header, ...rows].map((r) => r.map(esc).join(",")).join("\n");
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
