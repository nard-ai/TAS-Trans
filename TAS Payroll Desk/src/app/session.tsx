import { createContext, useContext, type ReactNode } from "react";
import type { Client, Cutoff, Db, User } from "../data/types";
import { useDb } from "../data/store";
import {
  canAdd,
  canDelete,
  canEdit,
  cutoffHalf,
  cutoffKey,
  isFmsClient,
  visibleClients,
} from "../data/rules";
import { isBilled } from "../data/api";

export interface SessionCtx {
  db: Db;
  user: User;
  agencyName: string;
  cutoff: Cutoff | null;
  cutoffClient: Client | null;
  cutoffKey: string | null;
  cutoffHalf: 1 | 2 | null;
  billed: boolean;
  fms: boolean;
  clients: Client[];
  can: { add: boolean; edit: boolean; del: boolean };
}

const Ctx = createContext<SessionCtx | null>(null);

export function SessionProvider({ children }: { children: ReactNode }) {
  const db = useDb();
  const user = db.users.find((u) => u.id === db.session?.userId);
  if (!user || !db.session) return null;
  const cutoff = db.session.cutoff;
  const cutoffClient = cutoff ? (db.clients.find((c) => c.id === cutoff.clientId) ?? null) : null;
  const value: SessionCtx = {
    db,
    user,
    agencyName: db.agencies.find((a) => a.code === db.session!.agencyCode)?.name ?? "",
    cutoff,
    cutoffClient,
    cutoffKey: cutoff ? cutoffKey(cutoff.end) : null,
    cutoffHalf: cutoff ? cutoffHalf(cutoff.end) : null,
    billed: isBilled(db, cutoff),
    fms: isFmsClient(cutoffClient),
    clients: visibleClients(db, user),
    can: { add: canAdd(user), edit: canEdit(user), del: canDelete(user) },
  };
  return <Ctx.Provider value={value}>{children}</Ctx.Provider>;
}

export function useSession() {
  const s = useContext(Ctx);
  if (!s) throw new Error("useSession outside SessionProvider");
  return s;
}
