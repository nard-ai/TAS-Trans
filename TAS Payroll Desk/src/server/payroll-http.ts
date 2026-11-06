import { gzipSync } from "node:zlib";
import type { Db } from "../data/types";
import { ApiError, handlers } from "./actions";
import { bind, snapshot } from "./desk";
import { connect, loadDb, saveDb } from "./persist";

function isDb(value: unknown): value is Db {
  if (!value || typeof value !== "object") return false;
  const row = value as Db;
  return Array.isArray(row.agencies) && Array.isArray(row.users) && Array.isArray(row.employees);
}

function safeMessage(err: unknown): string {
  const raw = err instanceof Error ? err.message : "The payroll database did not answer.";
  return raw.replace(/postgres(?:ql)?:\/\/\S+/gi, "postgresql://redacted");
}

let gate: Promise<unknown> = Promise.resolve();
function locked<T>(fn: () => Promise<T>): Promise<T> {
  const run = gate.then(fn, fn);
  gate = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

function database() {
  const url = process.env.DATABASE_URL;
  if (!url) throw new Error("DATABASE_URL is not set.");
  return connect(url);
}

export function writeJson(
  res: { statusCode: number; setHeader: (name: string, value: string) => void; end: (body?: Buffer) => void },
  out: { status: number; json: unknown },
) {
  const gz = gzipSync(Buffer.from(JSON.stringify(out.json)));
  res.statusCode = out.status;
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Content-Encoding", "gzip");
  res.setHeader("Cache-Control", "no-store");
  res.end(gz);
}

export async function payrollApi(
  kind: "db" | "call",
  method: string,
  body: unknown,
): Promise<{ status: number; json: unknown }> {
  try {
    const sql = database();
    if (kind === "call") {
      if (method !== "POST") return { status: 405, json: {} };
      const call = body as { name?: string; args?: unknown[] } | null;
      const fn = handlers[call?.name as keyof typeof handlers];
      if (!fn) return { status: 400, json: { error: "Unknown payroll action." } };
      try {
        const payload = await locked(async () => {
          const existing = await loadDb(sql);
          if (!existing) throw new Error("The payroll database is empty.");
          bind(existing);
          const result = await (fn as (...args: unknown[]) => Promise<unknown>)(...(call?.args ?? []));
          const saved = snapshot();
          await saveDb(sql, saved);
          return { result, db: saved };
        });
        return { status: 200, json: payload };
      } catch (err) {
        if (err instanceof ApiError) return { status: 400, json: { error: err.message, field: err.field } };
        throw err;
      }
    }
    if (method === "GET") {
      const db = await locked(async () => {
        const existing = await loadDb(sql);
        if (!existing) throw new Error("The payroll database is empty.");
        return existing;
      });
      return { status: 200, json: { db } };
    }
    if (method === "PUT") {
      if (!isDb(body)) return { status: 400, json: { error: "The payroll save was not a database." } };
      await locked(() => saveDb(sql, body));
      return { status: 200, json: { ok: true } };
    }
    return { status: 405, json: {} };
  } catch (err) {
    return { status: 500, json: { error: safeMessage(err) } };
  }
}
