import { payrollApi, writeJson } from "../src/server/payroll-http";

type NodeReq = {
  method?: string;
  body?: unknown;
};

type NodeRes = {
  statusCode: number;
  setHeader: (name: string, value: string) => void;
  end: (body?: Buffer) => void;
};

function parsed(body: unknown): unknown {
  if (body == null || body === "") return null;
  if (typeof body === "string") return JSON.parse(body);
  if (Buffer.isBuffer(body)) return JSON.parse(body.toString("utf8"));
  return body;
}

export default async function handler(req: NodeReq, res: NodeRes) {
  writeJson(res, await payrollApi("db", req.method || "GET", parsed(req.body)));
}
