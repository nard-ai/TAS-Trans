import { build } from "vite";
import fs from "node:fs";
import path from "node:path";

const root = process.cwd();
const out = path.join(root, ".vercel", "output");
fs.rmSync(out, { recursive: true, force: true });

await build({
  root,
  configFile: path.join(root, "vite.config.ts"),
  build: { outDir: path.join(out, "static"), emptyOutDir: true },
  logLevel: "error",
});

const vc = {
  runtime: "nodejs26.x",
  handler: "index.js",
  launcherType: "Nodejs",
  shouldAddHelpers: true,
  maxDuration: 60,
};

async function bundle(entry, name) {
  const dir = path.join(out, "functions", name);
  await build({
    configFile: false,
    root,
    logLevel: "error",
    build: {
      ssr: path.join(root, entry),
      outDir: dir,
      emptyOutDir: true,
      rollupOptions: {
        output: { format: "cjs", entryFileNames: "index.js", codeSplitting: false },
      },
    },
    ssr: { noExternal: true, external: ["node:fs", "node:path", "node:zlib", "node:stream", "node:buffer", "node:url", "node:crypto", "node:http", "node:https", "node:net", "node:tls"] },
  });
  fs.writeFileSync(path.join(dir, ".vc-config.json"), JSON.stringify(vc));
}

await bundle("scripts/fn-db.ts", "api/db.func");
await bundle("scripts/fn-call.ts", "api/call.func");

fs.writeFileSync(
  path.join(out, "config.json"),
  JSON.stringify({
    version: 3,
    routes: [
      { src: "/api/(.*)", dest: "/api/$1" },
      { handle: "filesystem" },
    ],
  }),
);

const db = fs.statSync(path.join(out, "functions", "api", "db.func", "index.js"));
console.log("vercel output ready", db.size);
