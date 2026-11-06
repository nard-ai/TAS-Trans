import { jsonResponse, payrollApi } from "../src/server/payroll-http";

export async function POST(request: Request) {
  const text = await request.text();
  const body = text ? JSON.parse(text) : null;
  return jsonResponse(await payrollApi("call", "POST", body));
}
