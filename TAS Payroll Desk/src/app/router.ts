import { useSyncExternalStore } from "react";

let previous = "home";

function read() {
  return window.location.hash.replace(/^#\/?/, "") || "home";
}
function subscribe(cb: () => void) {
  window.addEventListener("hashchange", cb);
  return () => window.removeEventListener("hashchange", cb);
}

export function useRoute() {
  const raw = useSyncExternalStore(subscribe, read);
  const [path, ...rest] = raw.split("/");
  return { path, param: rest.join("/") || undefined, raw };
}

export function navigate(to: string) {
  const cur = read();
  if (cur !== to) {
    if (!cur.startsWith("change-password")) previous = cur;
    window.location.hash = `/${to}`;
  }
}
export const previousRoute = () => previous;
