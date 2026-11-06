import { serve } from "../src/server/payroll-http";

type Req = { method?: string; url?: string; body?: unknown; text?: () => Promise<string> };

export const config = { runtime: "nodejs", maxDuration: 60 };

export default function handler(req: Req, res?: unknown) {
  return serve("db", req, res);
}

export function GET(req: Req) {
  return serve("db", req);
}

export function PUT(req: Req) {
  return serve("db", req);
}
