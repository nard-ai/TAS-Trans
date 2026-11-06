export function GET() {
  return new Response("pong", { headers: { "Cache-Control": "no-store" } });
}

export default function handler(_req: unknown, res?: { statusCode: number; setHeader: (name: string, value: string) => void; end: (body?: string) => void }) {
  if (res && typeof res.end === "function") {
    res.statusCode = 200;
    res.setHeader("Content-Type", "text/plain");
    res.end("pong");
    return;
  }
  return new Response("pong", { headers: { "Cache-Control": "no-store" } });
}
