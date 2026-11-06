import type { SVGProps } from "react";

type P = SVGProps<SVGSVGElement>;
const base = (p: P) => ({
  width: 16,
  height: 16,
  viewBox: "0 0 16 16",
  fill: "none",
  stroke: "currentColor",
  strokeWidth: 1.6,
  strokeLinecap: "round" as const,
  strokeLinejoin: "round" as const,
  "aria-hidden": true,
  ...p,
});

export const Icon = {
  search: (p: P) => (
    <svg {...base(p)}>
      <circle cx="7" cy="7" r="4.5" />
      <path d="m10.5 10.5 3 3" />
    </svg>
  ),
  check: (p: P) => (
    <svg {...base(p)}>
      <path d="m3.5 8.5 3 3 6-7" />
    </svg>
  ),
  alert: (p: P) => (
    <svg {...base(p)}>
      <circle cx="8" cy="8" r="6.25" />
      <path d="M8 4.75v3.75M8 11h.01" />
    </svg>
  ),
  info: (p: P) => (
    <svg {...base(p)}>
      <circle cx="8" cy="8" r="6.25" />
      <path d="M8 7.25V11M8 5h.01" />
    </svg>
  ),
  success: (p: P) => (
    <svg {...base(p)}>
      <circle cx="8" cy="8" r="6.25" />
      <path d="m5.5 8.25 1.75 1.75 3.25-3.75" />
    </svg>
  ),
  x: (p: P) => (
    <svg {...base(p)}>
      <path d="m4 4 8 8M12 4l-8 8" />
    </svg>
  ),
  chevronDown: (p: P) => (
    <svg {...base(p)}>
      <path d="m4.5 6.25 3.5 3.5 3.5-3.5" />
    </svg>
  ),
  chevronRight: (p: P) => (
    <svg {...base(p)}>
      <path d="m6.25 4.5 3.5 3.5-3.5 3.5" />
    </svg>
  ),
  arrowLeft: (p: P) => (
    <svg {...base(p)}>
      <path d="M12.5 8h-9M7 4.5 3.5 8 7 11.5" />
    </svg>
  ),
  lock: (p: P) => (
    <svg {...base(p)}>
      <rect x="3.5" y="7" width="9" height="6.5" rx="1" />
      <path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" />
    </svg>
  ),
  printer: (p: P) => (
    <svg {...base(p)}>
      <path d="M4.5 6V2.5h7V6M4.5 11.5h-2v-5h11v5h-2" />
      <rect x="4.5" y="9.5" width="7" height="4" />
    </svg>
  ),
  plus: (p: P) => (
    <svg {...base(p)}>
      <path d="M8 3.5v9M3.5 8h9" />
    </svg>
  ),
  menu: (p: P) => (
    <svg {...base(p)}>
      <path d="M3 4.5h10M3 8h10M3 11.5h10" />
    </svg>
  ),
};
