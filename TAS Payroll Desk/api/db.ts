import { jsonResponse, payrollApi, payrollProbe } from "../src/server/payroll-http";

export async function GET(request: Request) {
  if (request.url.includes("probe=1")) {
    return jsonResponse(await payrollProbe());
  }
  return jsonResponse(await payrollApi("db", "GET", null));
}

export async function PUT(request: Request) {
  const text = await request.text();
  const body = text ? JSON.parse(text) : null;
  return jsonResponse(await payrollApi("db", "PUT", body));
}
