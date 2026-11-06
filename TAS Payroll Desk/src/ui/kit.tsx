import {
  createContext,
  forwardRef,
  useCallback,
  useContext,
  useEffect,
  useId,
  useRef,
  useState,
  type ButtonHTMLAttributes,
  type InputHTMLAttributes,
  type ReactNode,
  type SelectHTMLAttributes,
  type TextareaHTMLAttributes,
} from "react";
import { Icon } from "./icons";

export function cx(...c: (string | false | null | undefined)[]) {
  return c.filter(Boolean).join(" ");
}

/* ---------- Buttons ----------
 * primary   yellow, one per screen, the step that commits work
 * secondary blue grey outline, every other action
 * quiet     blue text, navigation and low-weight actions
 * danger    red outline, destructive actions (confirm first)
 * dangerSolid / confirm  only inside confirmation dialogs
 */
export type BtnVariant = "primary" | "secondary" | "quiet" | "danger" | "dangerSolid" | "confirm";
const btnBase =
  "inline-flex items-center justify-center gap-1.5 h-8 px-3 rounded-ctl border text-sm font-medium whitespace-nowrap transition-colors select-none disabled:cursor-not-allowed";
const btnVariants: Record<BtnVariant, string> = {
  primary:
    "bg-yellow-500 border-yellow-600 text-grey-900 hover:bg-yellow-400 active:bg-yellow-600 disabled:bg-grey-100 disabled:border-grey-100 disabled:text-grey-400",
  secondary:
    "bg-white border-grey-200 text-grey-800 hover:bg-grey-50 hover:border-grey-300 active:bg-grey-100 disabled:bg-white disabled:border-grey-100 disabled:text-grey-300",
  quiet:
    "bg-transparent border-transparent text-blue-700 hover:bg-blue-50 hover:text-blue-800 disabled:text-grey-300 disabled:bg-transparent px-2",
  danger:
    "bg-white border-red-200 text-red-500 hover:bg-red-50 hover:border-red-300 disabled:bg-white disabled:border-grey-100 disabled:text-grey-300",
  dangerSolid:
    "bg-red-500 border-red-600 text-white hover:bg-red-600 disabled:bg-grey-100 disabled:border-grey-100 disabled:text-grey-400",
  confirm:
    "bg-blue-700 border-blue-800 text-white hover:bg-blue-800 disabled:bg-grey-100 disabled:border-grey-100 disabled:text-grey-400",
};

export const Spinner = ({ className }: { className?: string }) => (
  <svg
    className={cx("animate-spin size-3.5", className)}
    viewBox="0 0 16 16"
    fill="none"
    aria-hidden
  >
    <circle cx="8" cy="8" r="6" stroke="currentColor" strokeOpacity="0.25" strokeWidth="2" />
    <path d="M14 8a6 6 0 0 0-6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
  </svg>
);

type ButtonProps = ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: BtnVariant;
  busy?: boolean;
  icon?: ReactNode;
};
export const Button = forwardRef<HTMLButtonElement, ButtonProps>(function Button(
  { variant = "secondary", busy, icon, className, children, disabled, type = "button", ...rest },
  ref,
) {
  return (
    <button
      ref={ref}
      type={type}
      disabled={disabled || busy}
      aria-busy={busy || undefined}
      className={cx(btnBase, btnVariants[variant], className)}
      {...rest}
    >
      {busy ? <Spinner /> : icon}
      {children}
    </button>
  );
});

/* ---------- Fields ---------- */

const FieldCtx = createContext<{ id?: string; invalid?: boolean; describedBy?: string }>({});

export function Field({
  label,
  hint,
  error,
  required,
  children,
  className,
  aside,
}: {
  label: ReactNode;
  hint?: ReactNode;
  error?: string | null;
  required?: boolean;
  children: ReactNode;
  className?: string;
  aside?: ReactNode;
}) {
  const id = useId();
  const msgId = `${id}-msg`;
  return (
    <FieldCtx.Provider
      value={{ id, invalid: !!error, describedBy: error || hint ? msgId : undefined }}
    >
      <div className={cx("min-w-0", className)}>
        <div className="mb-1 flex items-baseline justify-between gap-2">
          <label htmlFor={id} className="text-xs font-medium text-grey-700">
            {label}
            {required && <span className="ml-0.5 text-red-500">*</span>}
          </label>
          {aside}
        </div>
        {children}
        {error ? (
          <p id={msgId} className="mt-1 flex items-start gap-1 text-xs font-medium text-red-600">
            <Icon.alert className="mt-px size-3.5 shrink-0" />
            {error}
          </p>
        ) : hint ? (
          <p id={msgId} className="mt-1 text-xs text-grey-500">
            {hint}
          </p>
        ) : null}
      </div>
    </FieldCtx.Provider>
  );
}

const inputBase =
  "h-8 w-full rounded-ctl border bg-white px-2.5 text-sm text-grey-900 placeholder:text-grey-400 transition-colors focus:outline-none focus:ring-2 disabled:bg-grey-50 disabled:text-grey-500 disabled:cursor-not-allowed read-only:bg-grey-50 read-only:border-grey-100 read-only:text-grey-800 read-only:focus:ring-0";
// <select> matches :read-only in Chromium, so selects use a base without read-only styles.
const selectBase =
  'h-8 w-full rounded-ctl border bg-white px-2.5 text-sm text-grey-900 transition-colors focus:outline-none focus:ring-2 disabled:bg-grey-50 disabled:text-grey-500 disabled:cursor-not-allowed'
const inputOk = "border-grey-200 hover:border-grey-300 focus:border-blue-500 focus:ring-blue-100";
const inputBad = "border-red-400 bg-red-50 focus:border-red-500 focus:ring-red-100";

export const TextInput = forwardRef<
  HTMLInputElement,
  InputHTMLAttributes<HTMLInputElement> & { invalid?: boolean }
>(function TextInput({ className, invalid, ...rest }, ref) {
  const f = useContext(FieldCtx);
  const bad = invalid ?? f.invalid;
  return (
    <input
      ref={ref}
      id={rest.id ?? f.id}
      aria-invalid={bad || undefined}
      aria-describedby={f.describedBy}
      className={cx(inputBase, bad ? inputBad : inputOk, className)}
      {...rest}
    />
  );
});

export function NumberInput({
  value,
  onValue,
  decimals = 2,
  className,
  invalid,
  ...rest
}: Omit<InputHTMLAttributes<HTMLInputElement>, "value" | "onChange"> & {
  value: number | null;
  onValue: (n: number | null) => void;
  decimals?: number;
  invalid?: boolean;
}) {
  const f = useContext(FieldCtx);
  const fmt = (v: number | null) =>
    v == null
      ? ""
      : decimals === 0
        ? String(v)
        : v.toLocaleString("en-PH", {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals,
          });
  const [text, setText] = useState(fmt(value));
  const [focus, setFocus] = useState(false);
  useEffect(() => {
    if (!focus) setText(fmt(value));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, focus]);
  const bad = invalid ?? f.invalid;
  return (
    <input
      id={rest.id ?? f.id}
      inputMode="decimal"
      aria-invalid={bad || undefined}
      aria-describedby={f.describedBy}
      className={cx(inputBase, bad ? inputBad : inputOk, "text-right", className)}
      value={text}
      onFocus={(e) => {
        setFocus(true);
        setText(value == null ? "" : String(value));
        requestAnimationFrame(() => e.target.select());
      }}
      onBlur={() => setFocus(false)}
      onChange={(e) => {
        setText(e.target.value);
        const n = parseFloat(e.target.value.replace(/,/g, ""));
        onValue(Number.isNaN(n) ? null : n);
      }}
      {...rest}
    />
  );
}

export function Select({
  className,
  invalid,
  children,
  ...rest
}: SelectHTMLAttributes<HTMLSelectElement> & { invalid?: boolean }) {
  const f = useContext(FieldCtx);
  const bad = invalid ?? f.invalid;
  return (
    <div className={cx("relative", className)}>
      <select
        id={rest.id ?? f.id}
        aria-invalid={bad || undefined}
        aria-describedby={f.describedBy}
        className={cx(
          inputBase,
          bad ? inputBad : inputOk,
          "appearance-none pr-8 disabled:opacity-100",
        )}
        {...rest}
      >
        {children}
      </select>
      <Icon.chevronDown className="pointer-events-none absolute right-2.5 top-1/2 size-4 -translate-y-1/2 text-grey-500" />
    </div>
  );
}

export function Textarea({ className, ...rest }: TextareaHTMLAttributes<HTMLTextAreaElement>) {
  const f = useContext(FieldCtx);
  return (
    <textarea
      id={rest.id ?? f.id}
      className={cx(inputBase, inputOk, "h-auto min-h-16 py-1.5", className)}
      {...rest}
    />
  );
}

export function Checkbox({
  checked,
  onChange,
  label,
  hint,
  disabled,
  className,
}: {
  checked: boolean;
  onChange: (v: boolean) => void;
  label: ReactNode;
  hint?: ReactNode;
  disabled?: boolean;
  className?: string;
}) {
  return (
    <label
      className={cx(
        "group flex items-start gap-2",
        disabled ? "cursor-not-allowed opacity-70" : "cursor-pointer",
        className,
      )}
    >
      <input
        type="checkbox"
        className="peer sr-only"
        checked={checked}
        disabled={disabled}
        onChange={(e) => onChange(e.target.checked)}
      />
      <span
        className={cx(
          "mt-px grid size-4 shrink-0 place-items-center rounded-tag border transition-colors peer-focus-visible:ring-2 peer-focus-visible:ring-blue-300",
          checked
            ? "border-blue-600 bg-blue-600 text-white"
            : "border-grey-300 bg-white group-hover:border-grey-400",
        )}
      >
        {checked && <Icon.check className="size-3" strokeWidth={2.2} />}
      </span>
      <span className="min-w-0">
        <span className="block text-sm text-grey-900">{label}</span>
        {hint && <span className="mt-0.5 block text-xs text-grey-500">{hint}</span>}
      </span>
    </label>
  );
}

/** Read-only result value, styled so it is clearly not an input. */
export function Readout({
  children,
  align = "left",
  strong,
}: {
  children: ReactNode;
  align?: "left" | "right";
  strong?: boolean;
}) {
  const f = useContext(FieldCtx);
  return (
    <output
      id={f.id}
      className={cx(
        "flex h-8 items-center rounded-ctl bg-grey-50 px-2.5 text-sm text-grey-900",
        align === "right" && "justify-end",
        strong && "font-semibold",
      )}
    >
      {children}
    </output>
  );
}

/* ---------- Layout ---------- */

export function PageHeader({
  section,
  title,
  description,
  actions,
}: {
  section?: string;
  title: string;
  description?: ReactNode;
  actions?: ReactNode;
}) {
  return (
    <header className="no-print mb-4 flex flex-wrap items-end justify-between gap-x-6 gap-y-3">
      <div className="min-w-0">
        {section && <p className="mb-0.5 text-xs font-medium text-grey-500">{section}</p>}
        <h1 className="text-xl font-semibold tracking-[-0.01em] text-grey-900">{title}</h1>
        {description && <p className="mt-1 max-w-[72ch] text-sm text-grey-600">{description}</p>}
      </div>
      {actions && <div className="flex flex-wrap items-center gap-2">{actions}</div>}
    </header>
  );
}

export function Panel({ children, className }: { children: ReactNode; className?: string }) {
  return (
    <section className={cx("min-w-0 rounded-panel border border-grey-100 bg-white", className)}>
      {children}
    </section>
  );
}

export function PanelHeader({
  title,
  meta,
  actions,
  className,
}: {
  title: ReactNode;
  meta?: ReactNode;
  actions?: ReactNode;
  className?: string;
}) {
  return (
    <div
      className={cx(
        "flex min-h-11 flex-wrap items-center justify-between gap-x-4 gap-y-2 border-b border-grey-100 px-4 py-2",
        className,
      )}
    >
      <div className="flex min-w-0 items-baseline gap-2">
        <h2 className="shrink-0 text-base font-semibold text-grey-900">{title}</h2>
        {meta && <span className="min-w-0 truncate text-xs text-grey-500">{meta}</span>}
      </div>
      {actions && <div className="flex flex-wrap items-center gap-2">{actions}</div>}
    </div>
  );
}

export function FormSection({
  title,
  description,
  children,
  className,
}: {
  title: string;
  description?: ReactNode;
  children: ReactNode;
  className?: string;
}) {
  return (
    <div className={cx("border-b border-grey-100 px-4 py-4 last:border-b-0", className)}>
      <div className="mb-3">
        <h3 className="text-sm font-semibold text-grey-900">{title}</h3>
        {description && <p className="mt-0.5 text-xs text-grey-500">{description}</p>}
      </div>
      {children}
    </div>
  );
}

export const Divider = () => <span className="mx-1 h-5 w-px bg-grey-100" aria-hidden />;

/* ---------- Status ---------- */

export type Tone = "error" | "success" | "info" | "neutral";
const noticeTone: Record<Tone, string> = {
  error: "border-red-400 bg-red-50 text-red-800",
  success: "border-cyan-600 bg-cyan-50 text-cyan-900",
  info: "border-blue-500 bg-blue-50 text-blue-900",
  neutral: "border-grey-300 bg-grey-50 text-grey-800",
};
const noticeIcon: Record<Tone, ReactNode> = {
  error: <Icon.alert className="size-4 text-red-500" />,
  success: <Icon.success className="size-4 text-cyan-700" />,
  info: <Icon.info className="size-4 text-blue-600" />,
  neutral: <Icon.info className="size-4 text-grey-500" />,
};
export function Notice({
  tone,
  title,
  children,
  actions,
  onDismiss,
  className,
}: {
  tone: Tone;
  title: ReactNode;
  children?: ReactNode;
  actions?: ReactNode;
  onDismiss?: () => void;
  className?: string;
}) {
  return (
    <div
      role={tone === "error" ? "alert" : "status"}
      className={cx(
        "flex items-start gap-2.5 rounded-ctl border-l-[3px] px-3 py-2.5",
        noticeTone[tone],
        className,
      )}
    >
      <span className="mt-px shrink-0">{noticeIcon[tone]}</span>
      <div className="min-w-0 flex-1">
        <p className="text-sm font-semibold">{title}</p>
        {children && <div className="mt-0.5 text-sm opacity-90">{children}</div>}
        {actions && <div className="mt-2 flex flex-wrap gap-2">{actions}</div>}
      </div>
      {onDismiss && (
        <button
          type="button"
          onClick={onDismiss}
          className="-m-1 rounded-ctl p-1 opacity-60 hover:opacity-100"
          aria-label="Dismiss"
        >
          <Icon.x className="size-3.5" />
        </button>
      )}
    </div>
  );
}

export type Status = { tone: Tone; title: string; body?: ReactNode } | null;
export function useStatus() {
  const [status, setStatus] = useState<Status>(null);
  const fail = useCallback((err: unknown, title?: string) => {
    const msg = err instanceof Error ? err.message : String(err);
    setStatus({ tone: "error", title: title ?? msg, body: title ? msg : undefined });
  }, []);
  return { status, setStatus, fail, clear: () => setStatus(null) };
}
export function StatusLine({
  status,
  onDismiss,
  className,
}: {
  status: Status;
  onDismiss?: () => void;
  className?: string;
}) {
  if (!status) return null;
  return (
    <Notice tone={status.tone} title={status.title} onDismiss={onDismiss} className={className}>
      {status.body}
    </Notice>
  );
}

export type TagTone = "neutral" | "blue" | "success" | "error" | "outline";
const tagTone: Record<TagTone, string> = {
  neutral: "bg-grey-100 text-grey-700",
  blue: "bg-blue-50 text-blue-800",
  success: "bg-cyan-50 text-cyan-800",
  error: "bg-red-50 text-red-600",
  outline: "border border-grey-200 text-grey-600",
};
export function Tag({
  tone = "neutral",
  children,
  className,
}: {
  tone?: TagTone;
  children: ReactNode;
  className?: string;
}) {
  return (
    <span
      className={cx(
        "inline-flex h-5 items-center whitespace-nowrap rounded-tag px-1.5 text-xs font-medium",
        tagTone[tone],
        className,
      )}
    >
      {children}
    </span>
  );
}

/* ---------- Tables ---------- */

export type Col<T> = {
  key: string;
  header: ReactNode;
  align?: "left" | "right" | "center";
  width?: string;
  render: (row: T, index: number) => ReactNode;
  className?: string;
};

export function DataTable<T>({
  columns,
  rows,
  rowKey,
  selected,
  onSelect,
  empty,
  maxHeight,
  foot,
  rowClass,
  dense,
  caption,
}: {
  columns: Col<T>[];
  rows: T[];
  rowKey: (row: T) => string;
  selected?: string | null;
  onSelect?: (row: T) => void;
  empty?: ReactNode;
  maxHeight?: string;
  foot?: ReactNode;
  rowClass?: (row: T) => string | false | undefined;
  dense?: boolean;
  caption?: string;
}) {
  const align = (a?: string) =>
    a === "right" ? "text-right" : a === "center" ? "text-center" : "text-left";
  return (
    <div className="overflow-auto" style={maxHeight ? { maxHeight } : undefined}>
      <table className="w-full border-separate border-spacing-0 text-sm">
        {caption && <caption className="sr-only">{caption}</caption>}
        <thead>
          <tr>
            {columns.map((c) => (
              <th
                key={c.key}
                scope="col"
                style={c.width ? { width: c.width } : undefined}
                className={cx(
                  "sticky top-0 z-[1] h-8 whitespace-nowrap border-b border-grey-100 bg-grey-50 text-xs font-semibold text-grey-600",
                  dense ? "px-2.5" : "px-3",
                  align(c.align),
                )}
              >
                {c.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && empty ? (
            <tr>
              <td colSpan={columns.length} className="p-0">
                {empty}
              </td>
            </tr>
          ) : (
            rows.map((r, i) => {
              const k = rowKey(r);
              const isSel = selected === k;
              return (
                <tr
                  key={k}
                  onClick={onSelect ? () => onSelect(r) : undefined}
                  onKeyDown={onSelect ? (e) => e.key === "Enter" && onSelect(r) : undefined}
                  tabIndex={onSelect ? 0 : undefined}
                  aria-selected={onSelect ? isSel : undefined}
                  className={cx(
                    "group",
                    onSelect && "cursor-pointer focus:outline-none focus-visible:bg-blue-50",
                    isSel
                      ? "bg-blue-50 shadow-[inset_3px_0_0_var(--color-blue-600)]"
                      : onSelect && "hover:bg-grey-50",
                    rowClass?.(r),
                  )}
                >
                  {columns.map((c) => (
                    <td
                      key={c.key}
                      className={cx(
                        dense ? "h-8 px-2.5" : "h-9 px-3",
                        "whitespace-nowrap border-b border-grey-100 align-middle",
                        align(c.align),
                        c.className,
                      )}
                    >
                      {c.render(r, i)}
                    </td>
                  ))}
                </tr>
              );
            })
          )}
        </tbody>
        {foot && <tfoot>{foot}</tfoot>}
      </table>
    </div>
  );
}

export function EmptyState({ children, action }: { children: ReactNode; action?: ReactNode }) {
  return (
    <div className="flex flex-col items-start gap-3 px-4 py-8 text-sm text-grey-600">
      <p>{children}</p>
      {action}
    </div>
  );
}

export function SearchInput({
  value,
  onChange,
  placeholder,
  className,
  label,
}: {
  value: string;
  onChange: (v: string) => void;
  placeholder: string;
  className?: string;
  label: string;
}) {
  return (
    <div className={cx("relative", className)}>
      <Icon.search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-grey-400" />
      <TextInput
        aria-label={label}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="pl-8"
      />
    </div>
  );
}

/* ---------- Tabs ---------- */

export function Tabs<K extends string>({
  tabs,
  value,
  onChange,
}: {
  tabs: { key: K; label: string; count?: number }[];
  value: K;
  onChange: (k: K) => void;
}) {
  return (
    <div
      role="tablist"
      className="flex gap-1 overflow-x-auto overflow-y-hidden border-b border-grey-100 px-2"
    >
      {tabs.map((t) => (
        <button
          key={t.key}
          role="tab"
          type="button"
          aria-selected={value === t.key}
          onClick={() => onChange(t.key)}
          className={cx(
            "-mb-px flex h-10 items-center gap-1.5 whitespace-nowrap border-b-2 px-2.5 text-sm font-medium transition-colors",
            value === t.key
              ? "border-blue-600 text-blue-800"
              : "border-transparent text-grey-600 hover:text-grey-900",
          )}
        >
          {t.label}
          {t.count != null && (
            <span className={cx("text-xs", value === t.key ? "text-blue-700" : "text-grey-400")}>
              {t.count}
            </span>
          )}
        </button>
      ))}
    </div>
  );
}

/* ---------- Definition rows ---------- */

export function KeyValues({
  items,
  cols = 1,
}: {
  items: { label: string; value: ReactNode; align?: "right" }[];
  cols?: 1 | 2 | 3 | 4;
}) {
  const grid = {
    1: "grid-cols-1",
    2: "sm:grid-cols-2",
    3: "sm:grid-cols-3",
    4: "sm:grid-cols-2 lg:grid-cols-4",
  }[cols];
  return (
    <dl className={cx("grid gap-x-6 gap-y-2.5", grid)}>
      {items.map((it) => (
        <div key={it.label} className="min-w-0">
          <dt className="text-xs text-grey-500">{it.label}</dt>
          <dd
            className={cx(
              "mt-0.5 truncate text-sm font-medium text-grey-900",
              it.align === "right" && "text-right",
            )}
          >
            {it.value || <span className="text-grey-400">Not set</span>}
          </dd>
        </div>
      ))}
    </dl>
  );
}

/* ---------- Dialogs ---------- */

export function Dialog({
  open,
  title,
  children,
  footer,
  onClose,
  width = 460,
}: {
  open: boolean;
  title: string;
  children: ReactNode;
  footer?: ReactNode;
  onClose: () => void;
  width?: number;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const prev = document.activeElement as HTMLElement | null;
    const t = setTimeout(
      () =>
        ref.current
          ?.querySelector<HTMLElement>("input,select,textarea,button:not([data-close])")
          ?.focus(),
      20,
    );
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", onKey);
    return () => {
      clearTimeout(t);
      window.removeEventListener("keydown", onKey);
      prev?.focus?.();
    };
  }, [open, onClose]);
  if (!open) return null;
  return (
    <div
      className="no-print fixed inset-0 z-50 grid place-items-center bg-grey-900/45 p-4"
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        ref={ref}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        className="w-full rounded-panel bg-white shadow-dialog"
        style={{ maxWidth: width }}
      >
        <div className="flex items-center justify-between border-b border-grey-100 px-5 py-3">
          <h2 className="text-base font-semibold text-grey-900">{title}</h2>
          <button
            type="button"
            data-close
            onClick={onClose}
            className="-mr-1 rounded-ctl p-1 text-grey-500 hover:bg-grey-50 hover:text-grey-800"
            aria-label="Close"
          >
            <Icon.x />
          </button>
        </div>
        <div className="px-5 py-4 text-sm text-grey-700">{children}</div>
        {footer && (
          <div className="flex justify-end gap-2 border-t border-grey-100 bg-grey-50 px-5 py-3 rounded-b-panel">
            {footer}
          </div>
        )}
      </div>
    </div>
  );
}

type ConfirmOpts = {
  title: string;
  body: ReactNode;
  confirmLabel: string;
  tone?: "danger" | "default";
  acknowledge?: string;
};
const ConfirmCtx = createContext<(o: ConfirmOpts) => Promise<boolean>>(async () => false);
export const useConfirm = () => useContext(ConfirmCtx);

export function ConfirmProvider({ children }: { children: ReactNode }) {
  const [opts, setOpts] = useState<ConfirmOpts | null>(null);
  const [ack, setAck] = useState(false);
  const resolver = useRef<(v: boolean) => void>(undefined);
  const confirm = useCallback((o: ConfirmOpts) => {
    setAck(false);
    setOpts(o);
    return new Promise<boolean>((res) => (resolver.current = res));
  }, []);
  const close = useCallback((v: boolean) => {
    resolver.current?.(v);
    setOpts(null);
  }, []);
  const cancel = useCallback(() => close(false), [close]);
  return (
    <ConfirmCtx.Provider value={confirm}>
      {children}
      <Dialog
        open={!!opts}
        title={opts?.title ?? ""}
        onClose={cancel}
        footer={
          <>
            <Button onClick={cancel}>Cancel</Button>
            <Button
              variant={opts?.tone === "danger" ? "dangerSolid" : "confirm"}
              disabled={!!opts?.acknowledge && !ack}
              onClick={() => close(true)}
            >
              {opts?.confirmLabel}
            </Button>
          </>
        }
      >
        <div className="space-y-3">
          <div className="leading-5">{opts?.body}</div>
          {opts?.acknowledge && (
            <Checkbox checked={ack} onChange={setAck} label={opts.acknowledge} />
          )}
        </div>
      </Dialog>
    </ConfirmCtx.Provider>
  );
}

/* ---------- Money and totals ---------- */

export function Amount({
  value,
  strong,
  muted,
  negative,
}: {
  value: number | null | undefined;
  strong?: boolean;
  muted?: boolean;
  negative?: boolean;
}) {
  if (value == null) return <span className="text-grey-400">—</span>;
  return (
    <span
      className={cx(
        "num whitespace-nowrap",
        strong && "font-semibold",
        muted || value === 0 ? "text-grey-400" : "text-grey-900",
        negative && value !== 0 && "text-red-600",
      )}
    >
      {value.toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
    </span>
  );
}
