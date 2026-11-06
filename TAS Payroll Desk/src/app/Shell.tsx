import { useEffect, type ReactNode } from "react";
import { useSession } from "./session";
import { navigate, useRoute } from "./router";
import { ROUTES, NAV_GROUPS, routeAllowed } from "./routes";
import { Button, Notice, Tag, cx, useConfirm } from "../ui/kit";
import { Icon } from "../ui/icons";
import { logout, setFlash } from "../data/api";
import { canSee } from "../data/rules";
import { fmtRange } from "../ui/format";

export function Brand({ compact }: { compact?: boolean }) {
  return (
    <div className="flex items-center gap-2.5">
      <div className="grid size-8 shrink-0 place-items-center rounded-ctl bg-blue-900 text-[11px] font-bold tracking-wide text-white">
        TAS
      </div>
      <div className="leading-tight">
        <div className="text-sm font-semibold text-grey-900">TAS Trans Group of Companies</div>
        {!compact && <div className="text-xs text-grey-600">HR and Payroll</div>}
      </div>
    </div>
  );
}

export function Shell({ children }: { children: ReactNode }) {
  const s = useSession();
  const { user, agencyName, db } = s;
  const route = useRoute();
  const confirm = useConfirm();

  useEffect(() => {
    window.scrollTo(0, 0);
  }, [route.raw]);

  const onLogout = async () => {
    const ok = await confirm({
      title: "Log out",
      body: "Log out of HR and Payroll? Your login lock is cleared and the logout time is recorded.",
      confirmLabel: "Log out",
    });
    if (!ok) return;
    await logout();
    navigate("home");
  };

  const groups = NAV_GROUPS.map((g) => ({
    group: g,
    items: ROUTES.filter((r) => r.group === g && routeAllowed(r, user)),
  })).filter((g) => g.items.length);

  const showCutoff = canSee(user, "payroll") || canSee(user, "timeKeeping");

  return (
    <div className="min-h-screen">
      <header className="no-print sticky top-0 z-30 flex h-14 items-center gap-4 border-b border-grey-100 bg-white px-4">
        <div className="flex items-center gap-2.5">
          <div className="grid size-8 shrink-0 place-items-center rounded-ctl bg-blue-900 text-[11px] font-bold tracking-wide text-white">
            TAS
          </div>
          <div className="leading-tight">
            <div className="text-sm font-semibold text-grey-900">TAS Trans Group of Companies</div>
            <div className="text-xs text-grey-600">
              HR and Payroll <span className="text-grey-300">·</span> {agencyName}
            </div>
          </div>
        </div>
        <div className="ml-auto flex items-center gap-1">
          <div className="mr-2 text-right leading-tight">
            <div className="text-sm font-medium text-grey-900">{user.name}</div>
            <div className="text-xs text-grey-500">
              {user.username} · {user.accountType}
              {user.superuser && " · Superuser"}
            </div>
          </div>
          <Button variant="quiet" onClick={() => navigate("change-password")} icon={<Icon.lock />}>
            Change password
          </Button>
          <Button variant="secondary" onClick={onLogout}>
            Log out
          </Button>
        </div>
      </header>

      <div className="flex">
        <nav
          aria-label="Main"
          className="no-print sticky top-14 h-[calc(100vh-3.5rem)] w-56 shrink-0 overflow-y-auto border-r border-grey-100 bg-white py-3"
        >
          {groups.map(({ group, items }) => (
            <div key={group} className="mb-3">
              {group !== "Work" && (
                <div className="px-4 pb-1 pt-2 text-xs font-medium text-grey-500">{group}</div>
              )}
              <ul>
                {items.map((r) => {
                  const href = r.href ?? r.path;
                  const active =
                    route.raw === href ||
                    (route.path === r.path && !r.href && !ROUTES.some((x) => x.href === route.raw));
                  return (
                    <li key={href}>
                      <a
                        href={`#/${href}`}
                        aria-current={active ? "page" : undefined}
                        className={cx(
                          "flex h-8 items-center px-4 text-sm transition-colors",
                          active
                            ? "bg-blue-50 font-medium text-blue-800 shadow-[inset_3px_0_0_var(--color-blue-600)]"
                            : "text-grey-700 hover:bg-grey-50 hover:text-grey-900",
                        )}
                      >
                        {r.label}
                      </a>
                    </li>
                  );
                })}
              </ul>
            </div>
          ))}
        </nav>

        <main className="min-w-0 flex-1">
          {showCutoff && <CutoffBar />}
          <div className="px-6 pb-16 pt-5">
            {db.flash && (
              <Notice
                tone={db.flash.tone}
                title={db.flash.title}
                onDismiss={() => setFlash(null)}
                className="mb-4"
              />
            )}
            {children}
          </div>
        </main>
      </div>
    </div>
  );
}

function CutoffBar() {
  const { cutoff, cutoffClient, cutoffKey, cutoffHalf, billed, fms } = useSession();
  return (
    <div className="no-print flex min-h-11 flex-wrap items-center gap-x-4 gap-y-1 border-b border-grey-100 bg-white/70 px-6 py-2 text-sm">
      <span className="text-xs font-medium text-grey-500">Selected cutoff</span>
      {cutoff && cutoffClient ? (
        <>
          <span className="font-medium text-grey-900">{cutoffClient.name}</span>
          <Tag tone={fms ? "blue" : "neutral"}>{fms ? "FMS" : "Staff"}</Tag>
          <span className="text-grey-300">|</span>
          <span className="text-grey-800">{fmtRange(cutoff.start, cutoff.end)}</span>
          <span className="text-grey-300">|</span>
          <span className="text-grey-800">Half {cutoffHalf}</span>
          <span className="text-grey-300">|</span>
          <span className="text-grey-600">
            DTR key <span className="font-medium text-grey-900">{cutoffKey}</span>
          </span>
          {billed ? (
            <Tag tone="success">
              <Icon.lock className="mr-1 size-3" />
              Billed, read-only
            </Tag>
          ) : (
            <Tag tone="outline">Open</Tag>
          )}
          <a
            href="#/cutoff"
            className="ml-auto text-sm font-medium text-blue-700 hover:text-blue-800"
          >
            Change cutoff
          </a>
        </>
      ) : (
        <>
          <span className="text-grey-700">None. Time keeping and payroll entry need a cutoff.</span>
          <a
            href="#/cutoff"
            className="ml-auto text-sm font-medium text-blue-700 hover:text-blue-800"
          >
            Select cutoff
          </a>
        </>
      )}
    </div>
  );
}
