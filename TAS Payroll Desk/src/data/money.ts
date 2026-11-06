import type { PayInputs } from "./types";

export const r2 = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;
export const uid = () => Math.random().toString(36).slice(2, 10);
export const emptyInputs = (): PayInputs => ({
  hours: {},
  allowance: 0,
  otherIncome: 0,
  taxRefund: 0,
});
