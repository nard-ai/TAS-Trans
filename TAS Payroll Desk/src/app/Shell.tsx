import { useEffect, useState, type ReactNode } from "react";
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
    <div className="flex min-w-0 items-center gap-2.5">
      <div className="grid size-8 shrink-0 place-items-center rounded-ctl bg-blue-900 text-[11px] font-bold tracking-wide text-white">
        TAS
      </div>
      <div className="min-w-0 leading-tight">
        <div className="truncate text-sm font-semibold text-grey-900">TAS Trans Group of Companies</div>
        {!compact && <div className="truncate text-xs text-grey-600">HR and Payroll</div>}
      </div>
    </div>
  );
}

export function Shell({ children }: { children: ReactNode }) {
  const s = useSession();
  const { user, agencyName, db } = s;
  const route = useRoute();
  const confirm = useConfirm();
  const [navOpen, setNavOpen] = useState(false);

  useEffect(() => {
    window.scrollTo(0, 0);
    setNavOpen(false);
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
      <header className="no-print sticky top-0 z-40 flex h-14 items-center gap-2 border-b border-grey-100 bg-white px-3 sm:gap-4 sm:px-4">
        <button
          type="button"
          className="grid size-8 shrink-0 place-items-center rounded-ctl text-grey-700 hover:bg-grey-50 md:hidden"
          aria-expanded={navOpen}
          aria-controls="main-nav"
          onClick={() => setNavOpen((open) => !open)}
        >
          <span className="sr-only">{navOpen ? "Close menu" : "Open menu"}</span>
          {navOpen ? <Icon.x /> : <Icon.menu />}
        </button>
        <div className="flex min-w-0 flex-1 items-center gap-2.5">
          <div className="grid size-8 shrink-0 place-items-center rounded-ctl bg-blue-900 text-[11px] font-bold tracking-wide text-white">
            TAS
          </div>
          <div className="min-w-0 leading-tight">
            <div className="truncate text-sm font-semibold text-grey-900">TAS Trans Group of Companies</div>
            <div className="truncate text-xs text-grey-600">
              <span className="hidden sm:inline">
                HR and Payroll <span className="text-grey-300">·</span>{" "}
              </span>
              {agencyName}
            </div>
          </div>
        </div>
        <div className="flex shrink-0 items-center gap-1">
          <div className="mr-2 hidden text-right leading-tight sm:block">
            <div className="text-sm font-medium text-grey-900">{user.name}</div>
            <div className="text-xs text-grey-500">
              {user.username} · {user.accountType}
              {user.superuser && " · Superuser"}
            </div>
          </div>
          <Button
            variant="quiet"
            onClick={() => navigate("change-password")}
            icon={<Icon.lock />}
            aria-label="Change password"
          >
            <span className="hidden sm:inline">Change password</span>
          </Button>
          <Button variant="secondary" onClick={onLogout}>
            Log out
          </Button>
        </div>
      </header>

      {navOpen && (
        <button
          type="button"
          className="no-print fixed inset-0 top-14 z-20 bg-grey-900/40 md:hidden"
          aria-label="Close menu"
          onClick={() => setNavOpen(false)}
        />
      )}

      <div className="flex">
        <nav
          id="main-nav"
          aria-label="Main"
          className={cx(
            "no-print top-14 z-30 h-[calc(100vh-3.5rem)] w-64 overflow-y-auto border-r border-grey-100 bg-white py-3 md:sticky md:z-auto md:block md:w-56 md:shrink-0",
            navOpen ? "fixed" : "hidden",
          )}
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
                          "flex h-10 items-center px-4 text-sm transition-colors md:h-8",
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
          <div className="px-4 pb-16 pt-5 sm:px-6">
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
    <div className="no-print flex min-h-11 flex-wrap items-center gap-x-4 gap-y-1 border-b border-grey-100 bg-white/70 px-4 py-2 text-sm sm:px-6">
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
            className="basis-full text-sm font-medium text-blue-700 hover:text-blue-800 sm:ml-auto sm:basis-auto"
          >
            Change cutoff
          </a>
        </>
      ) : (
        <>
          <span className="text-grey-700">None. Time keeping and payroll entry need a cutoff.</span>
          <a
            href="#/cutoff"
            className="basis-full text-sm font-medium text-blue-700 hover:text-blue-800 sm:ml-auto sm:basis-auto"
          >
            Select cutoff
          </a>
        </>
      )}
    </div>
  );
}
