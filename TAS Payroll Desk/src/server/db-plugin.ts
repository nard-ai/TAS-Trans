import type { IncomingMessage, ServerResponse } from "node:http";
import { loadEnv, type Plugin } from "vite";
import type { Db } from "../data/types";
import { seed } from "../data/seed";
import { ApiError, handlers } from "./actions";
import { bind, snapshot } from "./desk";
import { connect, ensureSchema, loadDb, saveDb, schemaPath } from "./persist";

function isDb(value: unknown): value is Db {
  if (!value || typeof value !== "object") return false;
  const row = value as Db;
  return Array.isArray(row.agencies) && Array.isArray(row.users) && Array.isArray(row.employees);
}

function safeMessage(err: unknown): string {
  const raw = err instanceof Error ? err.message : "The payroll database did not answer.";
  return raw.replace(/postgres(?:ql)?:\/\/\S+/gi, "postgresql://redacted");
}

function send(res: ServerResponse, status: number, body: unknown) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(body));
}

async function readBody(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    size += buf.length;
    if (size > 8_000_000) throw new Error("The payroll save is too large.");
    chunks.push(buf);
  }
  return Buffer.concat(chunks).toString("utf8");
}

export function payrollDbPlugin(root: string): Plugin {
  return {
    name: "payroll-db",
    configureServer(server) {
      const env = loadEnv(server.config.mode, root, "");
      const url = env.DATABASE_URL;
      let gate: Promise<unknown> = Promise.resolve();
      const locked = <T>(fn: () => Promise<T>): Promise<T> => {
        const run = gate.then(fn, fn);
        gate = run.then(
          () => undefined,
          () => undefined,
        );
        return run;
      };

      const ready = url
        ? ensureSchema(connect(url), schemaPath(root))
        : Promise.reject(new Error("DATABASE_URL is not set."));
      gate = ready.then(
        () => undefined,
        () => undefined,
      );

      server.middlewares.use(async (req, res, next) => {
        const pathname = req.url?.split("?")[0];
        if (pathname !== "/api/db" && pathname !== "/api/call") return next();
        try {
          await ready;
          if (!url) throw new Error("DATABASE_URL is not set.");
          const sql = connect(url);
          if (pathname === "/api/call") {
            if (req.method !== "POST") {
              res.statusCode = 405;
              res.end();
              return;
            }
            const body = JSON.parse(await readBody(req)) as { name?: string; args?: unknown[] };
            const fn = handlers[body.name as keyof typeof handlers];
            if (!fn) {
              send(res, 400, { error: "Unknown payroll action." });
              return;
            }
            try {
              const payload = await locked(async () => {
                const existing = await loadDb(sql);
                if (!existing) throw new Error("The payroll database is empty.");
                bind(existing);
                const result = await (fn as (...args: unknown[]) => Promise<unknown>)(
                  ...(body.args ?? []),
                );
                const saved = snapshot();
                await saveDb(sql, saved);
                return { result, db: saved };
              });
              send(res, 200, payload);
            } catch (err) {
              if (err instanceof ApiError) {
                send(res, 400, { error: err.message, field: err.field });
                return;
              }
              throw err;
            }
            return;
          }
          if (req.method === "GET") {
            const db = await locked(async () => {
              const existing = await loadDb(sql);
              if (existing) return existing;
              const seeded = seed();
              await saveDb(sql, seeded);
              return seeded;
            });
            send(res, 200, { db });
            return;
          }
          if (req.method === "PUT") {
            const parsed: unknown = JSON.parse(await readBody(req));
            if (!isDb(parsed)) {
              send(res, 400, { error: "The payroll save was not a database." });
              return;
            }
            await locked(() => saveDb(sql, parsed));
            send(res, 200, { ok: true });
            return;
          }
          res.statusCode = 405;
          res.end();
        } catch (err) {
          send(res, 500, { error: safeMessage(err) });
        }
      });
    },
  };
}
