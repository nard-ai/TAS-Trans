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

export async function payrollProbe(): Promise<{ status: number; json: unknown }> {
  try {
    const sql = database();
    const rows = await sql`SELECT count(*)::int AS n FROM employees`;
    const n = rows[0] ? Number((rows[0] as { n: number }).n) : 0;
    return { status: 200, json: { ok: true, employees: n } };
  } catch (err) {
    return { status: 200, json: { ok: false, error: safeMessage(err) } };
  }
}

type Incoming = {
  method?: string;
  url?: string;
  body?: unknown;
  text?: () => Promise<string>;
};

type Outgoing = {
  statusCode: number;
  setHeader: (name: string, value: string) => void;
  write: (chunk: Uint8Array) => void;
  end: (chunk?: Uint8Array | string) => void;
};

function isOutgoing(value: unknown): value is Outgoing {
  return Boolean(value) && typeof (value as Outgoing).setHeader === "function" && typeof (value as Outgoing).end === "function";
}

async function readBody(req: Incoming, method: string): Promise<unknown> {
  if (method === "GET" || method === "HEAD") return null;
  if (typeof req.text === "function") {
    const text = await req.text();
    return text ? JSON.parse(text) : null;
  }
  if (typeof req.body === "string") return req.body ? JSON.parse(req.body) : null;
  if (Buffer.isBuffer(req.body)) return JSON.parse(req.body.toString("utf8"));
  return req.body ?? null;
}

function writeOut(res: Outgoing, out: { status: number; json: unknown }) {
  const bytes = Buffer.from(JSON.stringify(out.json));
  res.statusCode = out.status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  const size = 512 * 1024;
  for (let i = 0; i < bytes.length; i += size) res.write(bytes.subarray(i, Math.min(i + size, bytes.length)));
  res.end();
}

export async function serve(kind: "db" | "call", req: Incoming, res?: unknown): Promise<Response | void> {
  const method = req?.method || (kind === "call" ? "POST" : "GET");
  try {
    const out = String(req?.url || "").includes("probe=1")
      ? await payrollProbe()
      : await payrollApi(kind, method, await readBody(req, method));
    if (isOutgoing(res)) {
      writeOut(res, out);
      return;
    }
    return jsonResponse(out);
  } catch (err) {
    const out = { status: 500, json: { error: safeMessage(err) } };
    if (isOutgoing(res)) {
      writeOut(res, out);
      return;
    }
    return jsonResponse(out);
  }
}

// shortcut: the full desk is about 10MB. Vercel rejects a buffered function body over 4.5MB, so this streams.
export function jsonResponse(out: { status: number; json: unknown }): Response {
  const bytes = Buffer.from(JSON.stringify(out.json));
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  };
  if (bytes.length <= 3_500_000) {
    return new Response(bytes, { status: out.status, headers });
  }
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const size = 512 * 1024;
      for (let i = 0; i < bytes.length; i += size) {
        controller.enqueue(bytes.subarray(i, Math.min(i + size, bytes.length)));
      }
      controller.close();
    },
  });
  return new Response(stream, { status: out.status, headers });
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
