import { serve } from "../src/server/payroll-http";

type Req = { method?: string; url?: string; body?: unknown; text?: () => Promise<string> };

export default function handler(req: Req, res?: unknown) {
  return serve("call", req, res);
}
