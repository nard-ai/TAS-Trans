import { serve } from "../src/server/payroll-http";

type Req = { method?: string; url?: string; body?: unknown; text?: () => Promise<string> };

export const config = { runtime: "nodejs", maxDuration: 60 };

export default function handler(req: Req, res?: unknown) {
  return serve("call", req, res);
}

export function POST(req: Req) {
  return serve("call", req);
}
