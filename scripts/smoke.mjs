#!/usr/bin/env node
// 离线冒烟测试：验证插件模块可加载、peer 依赖可解析、pdf_to_word 工具可注册。
// 不真实执行转换 / LLM 调用。
//
//   node scripts/smoke.mjs [dsh-cli-js]
// 缺省 dsh CLI 入口：自动探测（Windows nvm/nvm4w 常见布局）。
import { createRequire } from "node:module";
import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const PKG_ROOT = join(here, "..");

/** 探测一个能 resolve @deepseek-ai/dsh-tools 的 JS 锚点（模拟宿主 process.argv[1]）。 */
function findAnchor(explicit) {
  if (explicit) return explicit;
  const candidates = [];
  // 1. node 同级全局：<node-dir>/node_modules/@deepseek-ai/dsh/lib/bin.js
  candidates.push(join(dirname(process.execPath), "node_modules", "@deepseek-ai", "dsh", "lib", "bin.js"));
  // 2. 常见全局根
  for (const base of [
    process.env.USERPROFILE && join(process.env.USERPROFILE, "AppData", "Local", "nvm"),
    "C:/nvm4w/nodejs"
  ]) {
    if (!base) continue;
    candidates.push(join(base, "node_modules", "@deepseek-ai", "dsh", "lib", "bin.js"));
  }
  for (const c of candidates) {
    if (!existsSync(c)) continue;
    try {
      createRequire(c).resolve("@deepseek-ai/dsh-tools");
      return c;
    } catch {
      /* keep looking */
    }
  }
  throw new Error("未找到可解析 dsh-tools 的锚点；请显式传入 dsh CLI js 路径，如：\n  node scripts/smoke.mjs C:\\...\\dsh\\lib\\bin.js");
}

process.argv[1] = findAnchor(process.argv[2]);

const plugin = await import(new URL("../lib/index.js", import.meta.url).href);
if (plugin.name !== "pdf-to-word") throw new Error("插件名不符：" + plugin.name);
console.log("plugin name :", plugin.name);
console.log("inject      :", (plugin.inject || []).join(", "));

const registered = [];
const ctx = {
  get: (name) => {
    if (name === "llm") return { listProviders: () => ["local"] };
    if (name === "sandboxPolicy") return { resolve: () => ({ __policy: true }) };
    return undefined;
  },
  tools: { register: (def) => registered.push(def) },
  shell: {
    resolve: (req) => req,
    run: async () => {
      throw new Error("smoke: shell 不应被调用");
    }
  },
  fs: {
    resolve: async (p) => ({ targetKey: p, displayPath: p }),
    stat: async () => null,
    readBytes: async () => new Uint8Array(),
    writeText: async () => {}
  }
};

plugin.apply(ctx, {});

if (registered.length !== 1) throw new Error("期望注册 1 个工具，实际 " + registered.length);
const t = registered[0];
if (t.name !== "pdf_to_word") throw new Error("工具名不符：" + t.name);
if (typeof t.execute !== "function") throw new Error("execute 缺失");
// defineTool 将平铺参数包装为 JSON Schema：{type, properties, required}
const propDef = t.parameters.properties || t.parameters;
const params = Object.keys(propDef);
console.log("tool name   :", t.name);
console.log("parameters  :", params.join(", "));
for (const want of ["pdfPath", "mode", "verify", "maxVerifyPages", "outPath"]) {
  if (!params.includes(want)) throw new Error("参数缺失：" + want);
}
// defineTool 会把逐参数 required:true 提升为顶层 required 数组
if (!(Array.isArray(t.parameters.required) && t.parameters.required.includes("pdfPath")))
  throw new Error("pdfPath 应为必填");
if (!t.output || !t.output.schema || typeof t.output.render !== "function") throw new Error("output 结构不完整");

// 渲染函数干跑（正常结果形态）
const rendered = t.output.render(
  { pdfPath: "x.pdf" },
  {
    ok: true,
    docx: "C:/x.docx",
    report: "C:/a/verify_report.md",
    renderer: "pil",
    n_pages: 2,
    n_scan: 0,
    overall: { match: true, avg: 0.9, verified: 2 },
    pages: [{ page: 1, mode: "digital", match: true, score: 0.9, issues: [], note: "" }],
    warnings: []
  }
);
if (!rendered || !rendered[0] || typeof rendered[0].text !== "string") throw new Error("render 输出异常");
console.log("\nrender preview:\n" + rendered[0].text);

console.log("\nSMOKE OK");
