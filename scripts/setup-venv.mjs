#!/usr/bin/env node
// Create the plugin's Python virtualenv and install pipeline/requirements.txt.
//
//   node scripts/setup-venv.mjs            full setup (core + scan/OCR deps)
//   node scripts/setup-venv.mjs --core     core deps only (digital mode; no Paddle)
//
// Idempotent: an existing .venv is reused. Requires Python 3.10-3.12 on PATH
// (the pipeline was built on 3.12; cp312 wheels are assumed).
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const isWin = process.platform === "win32";
const venv = join(root, ".venv");
const pythonExe = isWin ? join(venv, "Scripts", "python.exe") : join(venv, "bin", "python");
const coreOnly = process.argv.includes("--core");

function run(cmd, args) {
  console.log(`$ ${cmd} ${args.join(" ")}`);
  const r = spawnSync(cmd, args, { stdio: "inherit" });
  if (r.status !== 0) {
    console.error(`Failed (exit ${r.status ?? "killed"}): ${cmd} ${args.join(" ")}`);
    process.exit(r.status ?? 1);
  }
}

// 1. Locate a base interpreter (Python 3.10-3.12).
const candidates = isWin ? ["python", "py", "python3"] : ["python3.12", "python3.11", "python3.10", "python3"];
let base = null;
for (const c of candidates) {
  const probe = spawnSync(c, ["--version"], { encoding: "utf8" });
  const out = `${probe.stdout ?? ""}${probe.stderr ?? ""}`;
  if (probe.status === 0 && /Python 3\.(1[0-2])\./.test(out)) {
    base = c;
    break;
  }
}
if (!base) {
  console.error("No Python 3.10-3.12 interpreter found on PATH. Install Python 3.12 and re-run.");
  process.exit(1);
}
console.log(`Using interpreter: ${base} (${(spawnSync(base, ["--version"], { encoding: "utf8" }).stdout ?? "").trim()})`);

// 2. Create the venv if missing.
if (!existsSync(pythonExe)) {
  run(base, ["-m", "venv", venv]);
}

// 3. Install dependencies.
let reqFile = join(root, "pipeline", "requirements.txt");
if (coreOnly) {
  const text = readFileSync(reqFile, "utf8");
  const core = text
    .split("\n")
    .filter((l) => {
      const t = l.trim();
      if (t.startsWith("#") || t === "") return true;
      return !/^(paddlepaddle|paddleocr)\b/i.test(t);
    })
    .join("\n");
  reqFile = join(root, "pipeline", "requirements-core.generated.txt");
  writeFileSync(reqFile, core + "\n");
  console.log("core mode: skipping Paddle (scan/OCR unavailable)");
}
run(pythonExe, ["-m", "pip", "install", "--upgrade", "pip"]);
run(pythonExe, ["-m", "pip", "install", "-r", reqFile]);

console.log(`\nDone. The plugin will use: ${pythonExe}`);
