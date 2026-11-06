import type { Db } from "../data/types";

let db: Db;

export function bind(next: Db) {
  db = next;
}

export function getDb(): Db {
  return db;
}

export function mutate(fn: (d: Db) => void) {
  fn(db);
  db = { ...db };
}

export function snapshot(): Db {
  return db;
}
