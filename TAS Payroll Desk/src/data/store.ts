import { useSyncExternalStore } from "react";
import type { Db } from "./types";

const blank = (): Db => ({
  agencies: [],
  users: [],
  company: null,
  clients: [],
  employees: [],
  shifts: [],
  holidays: [],
  leaves: [],
  overtimes: [],
  punches: [],
  staffFinal: {},
  fmsFinal: {},
  payrollAttendance: {},
  cutoffs: [],
  deductions: [],
  openPayroll: [],
  history: [],
  trips: [],
  payments: [],
  thirteenth: [],
  banks: [],
  branches: [],
  departments: [],
  sss: [],
  philhealth: [],
  hdmf: [],
  tax: [],
  taxCodes: [],
  logs: [],
  benefits: [],
  assignments: [],
  trainings: [],
  offenses: [],
  leaveBalances: [],
  appraisals: [],
  session: null,
  flash: null,
});

type Status = "loading" | "ready" | "error";

let db: Db = blank();
let status: Status = "loading";
let errorText = "The payroll database did not answer.";
let saveError: string | null = null;
const listeners = new Set<() => void>();

function emit() {
  listeners.forEach((l) => l());
}

async function readError(res: Response): Promise<string> {
  try {
    const body = (await res.json()) as { error?: string };
    return body.error || "The payroll database did not answer.";
  } catch {
    return "The payroll database did not answer.";
  }
}

export function replaceDb(next: Db) {
  db = next;
  status = "ready";
  emit();
}

export function noteSaveError(message: string | null) {
  saveError = message;
  emit();
}

async function boot() {
  try {
    const res = await fetch("/api/db");
    if (!res.ok) throw new Error(await readError(res));
    const body = (await res.json()) as { db: Db };
    db = body.db;
    status = "ready";
  } catch (err) {
    status = "error";
    errorText = err instanceof Error ? err.message : errorText;
  }
  emit();
}

void boot();

export const getDb = () => db;

function subscribe(l: () => void) {
  listeners.add(l);
  return () => listeners.delete(l);
}

export function useDb(): Db {
  return useSyncExternalStore(subscribe, getDb);
}

export function useDbStatus(): Status {
  return useSyncExternalStore(subscribe, () => status);
}

export function useDbError(): string {
  return useSyncExternalStore(subscribe, () => errorText);
}

export function useSaveError(): string | null {
  return useSyncExternalStore(subscribe, () => saveError);
}
