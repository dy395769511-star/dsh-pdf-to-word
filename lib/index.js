// lib/index.js — 静态 Cordis 插件入口：pdf-to-word
//
// 注册一个模型工具 `pdf_to_word`：
//   将 PDF 转换为 Word .docx（保真字体/表格/图片/边框/背景，扫描页走 OCR），
//   并可选用多模态大模型逐页做保真校验，写出校验报告。
//
// 实现方式：
//   - 转换由包内 Python 管线完成（pipeline/convert.py；扫描页调用
//     scanocr.py，预览渲染由 preview.py 承担）。通过 shell 服务按会话
//     解析沙箱策略后启动 venv 解释器，stdout 最后一行
//     `PDF2WORD_JSON <json>` 即结构化结果。
//   - 解释器解析顺序：行配置 python > PDF2WORD_PYTHON > 包内 .venv；
//     均未命中时首次调用自动运行 scripts/setup-venv.mjs 建立完整 .venv
//     （行配置 autoSetup: false 可禁用；结果 warnings 会提示自动创建耗时）。
//   - 校验走宿主 llm 服务（ctx.llm.stream）+ attachments 服务（图片落
//     库为 attachment 引用），provider/model 由行配置给出，默认
//     local / Qwen3.8-27B（凭据由 dsh 设置管理，不进入本包）。
//   - 报告通过 fs 服务写到 <pdf 所在目录>/pdf2w_assets/verify_report.md。
//   - 校验发现 medium+ 表格结构问题后可自动进入 verify→fix 闭环
//     （maxFixRounds，默认 2，上限 3）：fixdocx.py dump 导出表格结构，
//     LLM 输出白名单修复动作方案（7 类表格操作），fixdocx.py apply 确定性
//     执行（逐操作校验 + 全局不变式门禁 + 整表回滚），重渲染并仅复检
//     受影响页，循环至通过或达到轮次上限；方案落盘 pdf2w_assets/fix_plan_rN.json。
//
// 依赖服务：
//   硬依赖 inject: tools（注册工具）、shell（启动子进程）、fs（路径/读写）
//   软依赖 ctx.get: llm、attachments（缺失时跳过校验并给出警告）、
//                  sandboxPolicy（缺失时按 shell 服务默认策略）

import { createRequire } from "node:module";
import { existsSync } from "node:fs";
import { homedir, platform } from "node:os";
import { dirname, isAbsolute, join, normalize } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const PKG_ROOT = join(here, "..");
const PIPELINE_DIR = join(PKG_ROOT, "pipeline");
const CONVERT_PY = join(PIPELINE_DIR, "convert.py");
const FIXDOCX_PY = join(PIPELINE_DIR, "fixdocx.py");
const PREVIEW_PY = join(PIPELINE_DIR, "preview.py");

// ---------------------------------------------------------------------------
// peer 依赖多锚点加载（与 dsh-word-template 同因：本地 link: 安装时
// pnpm 建 junction，Node 按真实路径解析，裸 import 可能失败）
// ---------------------------------------------------------------------------
async function loadPeer(spec, check) {
	const attempts = [];
	try {
		const mod = await import(spec);
		if (check(mod)) return mod;
		attempts.push("self: 导出缺失");
	} catch (error) {
		attempts.push("self: " + (error && error.message ? error.message : String(error)));
	}
	const tryAnchor = async (label, anchorPath) => {
		try {
			const entry = createRequire(anchorPath).resolve(spec);
			const mod = await import(pathToFileURL(entry).href);
			if (check(mod)) return mod;
			attempts.push(label + ": 导出缺失");
			return null;
		} catch (error) {
			attempts.push(label + ": " + (error && error.message ? error.message : String(error)));
			return null;
		}
	};
	const main = process.argv[1];
	if (main && isAbsolute(main)) {
		const mod = await tryAnchor("host", main);
		if (mod) return mod;
	}
	const dshHome = process.env.DSH_HOME || join(homedir(), ".dsh");
	const mod = await tryAnchor("fallback", join(dshHome, "profiles", "node_modules", ".dsh-anchor.js"));
	if (mod) return mod;
	throw new Error(
		"无法解析 " + spec + "（peer 依赖，应由 DSH 宿主提供）。尝试过：" + attempts.join(" | ")
	);
}

const dshTools = await loadPeer("@deepseek-ai/dsh-tools", (m) => m && typeof m.defineTool === "function");
const dshLlm = await loadPeer("@deepseek-ai/dsh-llm", (m) => m && typeof m.createMessage === "function");
const { defineTool } = dshTools;
const { createMessage } = dshLlm;

/** Cordis 插件名（loader 诊断用）。 */
export const name = "pdf-to-word";
/** 硬依赖服务：tools（注册工具）、shell（启动 Python 管线）、fs（路径/文件）。 */
export const inject = ["tools", "shell", "fs"];

// ---------------------------------------------------------------------------
// 基础工具函数
// ---------------------------------------------------------------------------
function tail(value, n) {
	if (value === undefined || value === null) return "";
	const s = String(value);
	return s.length > n ? "…" + s.slice(-n) : s;
}

function sessionOf(exec) {
	try {
		return exec && exec.agent && exec.agent.session ? exec.agent.session : null;
	} catch {
		return null;
	}
}

function sessionCwd(exec) {
	const session = sessionOf(exec);
	try {
		const cwd = session && session.header && session.header.cwd;
		if (cwd) return cwd;
	} catch {
		/* ignore */
	}
	return null;
}

/** Python 解释器解析：行配置 python > 环境变量 PDF2WORD_PYTHON > <包内>/.venv。 */
function resolvePython(config) {
	if (config && typeof config.python === "string" && config.python.trim()) return config.python.trim();
	const env = process.env.PDF2WORD_PYTHON;
	if (env && env.trim()) return env.trim();
	const win = platform() === "win32";
	const venvPy = join(PKG_ROOT, ".venv", win ? "Scripts" : "bin", win ? "python.exe" : "python");
	if (existsSync(venvPy)) return venvPy;
	return null;
}

function round2(x) {
	return Math.round(x * 100) / 100;
}

/** 从 stdout 中提取最后一个 `PDF2WORD_JSON <json>` 行并解析。 */
function parsePipelineJson(stdout) {
	const MARK = "PDF2WORD_JSON";
	let last = null;
	for (const line of String(stdout).split(/\r?\n/)) {
		const i = line.indexOf(MARK);
		if (i === -1) continue;
		const body = line.slice(i + MARK.length).trim();
		if (body) last = body;
	}
	if (!last) return null;
	try {
		return JSON.parse(last);
	} catch (error) {
		throw new Error("无法解析管线输出 PDF2WORD_JSON：" + (error && error.message ? error.message : error));
	}
}

/** 从 stdout 中提取最后一个 `FIXDOCX_JSON <json>` 行并解析。 */
function parseFixJson(stdout) {
	const MARK = "FIXDOCX_JSON";
	let last = null;
	for (const line of String(stdout).split(/\r?\n/)) {
		const i = line.indexOf(MARK);
		if (i === -1) continue;
		const body = line.slice(i + MARK.length).trim();
		if (body) last = body;
	}
	if (!last) return null;
	try {
		return JSON.parse(last);
	} catch (error) {
		throw new Error("无法解析修复脚本输出 FIXDOCX_JSON：" + (error && error.message ? error.message : error));
	}
}

// ---------------------------------------------------------------------------
// verify→fix 闭环
// ---------------------------------------------------------------------------
/** 修复动作白名单及各操作必填字段。
 *  t=顶层表格索引；row/col 为 0 基视觉顺序；span 为跨列数。 */
const FIX_OPS = {
	setGridSpan: [["t", "int"], ["row", "int"], ["col", "int"], ["span", "intPos"]],
	setVMerge: [["t", "int"], ["row", "int"], ["col", "int"], ["merge", "enum"]],
	setCellText: [["t", "int"], ["row", "int"], ["col", "int"], ["text", "str"]],
	insertCell: [["t", "int"], ["row", "int"], ["col", "int"], ["text", "strOpt"], ["span", "intPosOpt"]],
	removeCell: [["t", "int"], ["row", "int"], ["col", "int"]],
	cloneRow: [["t", "int"], ["afterRow", "int"], ["sourceRow", "int"]],
	removeRow: [["t", "int"], ["row", "int"]]
};

/** 修复方案 prompt：严格 JSON 输出 + 动作白名单 + 硬约束。 */
function planFixPrompt({ dumpText, pageIssues, worstPages }) {
	return [
		"你是一个 Word .docx 表格结构修复工程师。下面是一个 PDF 转换生成的 docx 的表格结构转储和校验发现的结构问题，请输出确定性修复方案。",
		"",
		"## 表格结构转储",
		"字段说明：t = 文档中表格索引；grid = 各网格列宽（dxa，数组长度即列数）；rows 按视觉行顺序、0 基；单元格 x = 文本（可能截断）、gs = 跨列数、vm = 纵向合并（0 无、1 restart、2 continue）、w = 单元格宽度。",
		dumpText,
		"",
		"## 校验发现的问题",
		...pageIssues,
		"",
		"## 附图",
		"附图依次为第 " + worstPages.join("、") + " 页的原始 PDF 与最新 Word 预览（相邻两张为一页）。以 PDF 图为准判定正确内容与结构。",
		"",
		"## 动作白名单（只允许以下 7 种）",
		'- setGridSpan {"t","row","col","span"}：把 row/col 单元格跨列数改为 span',
		'- setVMerge {"t","row","col","merge"}：纵向合并模式，merge ∈ restart|continue|none',
		'- setCellText {"t","row","col","text"}：替换单元格文本（保留原字体样式）',
		'- insertCell {"t","row","col","text?","span?"}：在当前 col 之前插入一个单元格（span 默认 1）',
		'- removeCell {"t","row","col"}：删除 row/col 单元格',
		'- cloneRow {"t","afterRow","sourceRow"}：在 afterRow 行后复制一份 sourceRow 行',
		'- removeRow {"t","row"}：整行删除第 row 行',
		"",
		"## 硬约束",
		"1. 所有动作应用后，每一行的 gridSpan 之和必须恰好等于 dump 中 grid 的列数（过程中可暂时不等，如先扩大某格跨列再收缩相邻格，但终态必须相等）。",
		"2. vMerge=continue 的单元格，其正上方同一网格列必须存在 restart 或 continue 单元格。",
		"3. 只修改与问题直接相关的单元格/行，不要触碰其它部分。",
		"4. 单元格文本必须取自 PDF 图中的原始内容；禁止编造、改写或删减 PDF 中不存在的内容。",
		"5. 字体、字号、图片、段落样式类问题不属于本次修复范围，不要输出对应动作。",
		"6. 单轮动作数建议不超过 40 个。",
		"7. 若没有可修复的表格结构问题，输出空方案。",
		"8. vMerge=continue 单元格的文本应为空：合并区文字一律写在 restart 单元格里；需要合并的列先用 setCellText 清空该格文本，再用 setVMerge 设置合并。",
		"9. 同一表格中语义相同的列（如子项列、要求列、结果列）在所有行中保持同一 col 索引。写入文本前，先对照同表相邻的同类行确认该列的 col 位置。",
		"",
		"只输出一个 JSON 对象，不要输出任何其他文字，格式：",
		'{"actions": [{"op": "动作名", ...}], "notes": "一句话方案说明"}'
	].join("\n");
}

/** 解析 LLM 修复方案：剥围栏 → 截取首个 { 到末个 } → JSON → 逐动作校验字段；
 *  非法动作丢弃并计入 warnings。 */
function parseFixPlan(raw, warnings) {
	let text = String(raw === undefined || raw === null ? "" : raw).trim();
	if (!text) {
		warnings.push("修复轮：模型方案输出为空");
		return [];
	}
	const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
	if (fence) text = fence[1].trim();
	const start = text.indexOf("{");
	const end = text.lastIndexOf("}");
	if (start === -1 || end <= start) {
		warnings.push("修复轮：模型方案中未找到 JSON");
		return [];
	}
	let data;
	try {
		data = JSON.parse(text.slice(start, end + 1));
	} catch {
		warnings.push("修复轮：模型方案不是合法 JSON");
		return [];
	}
	const list = Array.isArray(data && data.actions) ? data.actions : [];
	const out = [];
	let dropped = 0;
	for (const a of list.slice(0, 200)) {
		const spec = a && typeof a === "object" && typeof a.op === "string" ? FIX_OPS[a.op] : null;
		if (!spec) {
			dropped += 1;
			continue;
		}
		let okAll = true;
		for (const [key, kind] of spec) {
			const v = a[key];
			if (kind === "int" || kind === "intPos") {
				if (!Number.isInteger(v) || v < 0 || (kind === "intPos" && v < 1)) { okAll = false; break; }
			} else if (kind === "enum") {
				if (v !== "restart" && v !== "continue" && v !== "none") { okAll = false; break; }
			} else if (kind === "str") {
				if (typeof v !== "string") { okAll = false; break; }
			} else if (kind === "strOpt") {
				if (v !== undefined && typeof v !== "string") { okAll = false; break; }
			} else if (kind === "intPosOpt") {
				if (v !== undefined && (!Number.isInteger(v) || v < 1)) { okAll = false; break; }
			}
		}
		if (okAll) out.push(a);
		else dropped += 1;
	}
	if (dropped > 0) warnings.push("修复轮：丢弃 " + dropped + " 个字段不合法的动作");
	return out;
}

/** 该页是否存在可修复的 medium+ 问题（校验流程性失败除外——不可修复）。 */
function pageHasFixableIssues(p) {
	return (Array.isArray(p.issues) ? p.issues : []).some(
		(i) => (i.severity === "medium" || i.severity === "high")
			&& !String(i.description || "").startsWith("校验失败")
	);
}

// ---------------------------------------------------------------------------
// 校验逻辑
// ---------------------------------------------------------------------------
/** 选页：必含首页与末页，其余按 表格*2+图片 降序补足，最后均匀化（页数少时）。 */
function pickPages(pages, max) {
	const list = (Array.isArray(pages) ? pages : [])
		.map((p) => ({
			page: Number(p && p.page) || 0,
			tables: Number(p && p.tables) || 0,
			images: Number(p && p.images) || 0
		}))
		.sort((a, b) => a.page - b.page);
	const n = list.length;
	if (n === 0) return [];
	if (n <= max) return list.map((p) => p.page);
	const chosen = new Set([list[0].page, list[n - 1].page]);
	const rest = list
		.filter((p) => !chosen.has(p.page))
		.sort((a, b) => (b.tables * 2 + b.images) - (a.tables * 2 + a.images) || a.page - b.page);
	for (const p of rest) {
		if (chosen.size >= max) break;
		chosen.add(p.page);
	}
	return [...chosen].sort((a, b) => a - b);
}

function verifyPrompt(pageNo) {
	return [
		"你是一个严格的文档排版比对审校员。下面给你两张图：图 1 是原始 PDF 第 " + pageNo + " 页的渲染图，图 2 是转换生成的 Word 文档第 " + pageNo + " 页的预览渲染图。请逐项比对并判定转换质量。",
		"",
		"校验要点：",
		"1. 文字内容：是否完整、无缺漏、无错字（极细微的换行差异可接受）。",
		"2. 字体与字号：标题/正文/代码的字号层级是否一致。",
		"3. 段落结构：段落划分、对齐、缩进是否一致。",
		"4. 表格（重点，逐项核对样式与结构）：",
		"   4a. 列数与列对齐：Word 表格的列数是否与 PDF 一致；每一行各单元格的文本是否落在与 PDF 相同的列。某列内容缺失、整行文字右移/左移错位、标签与其值被拆到不同位置，均判为问题。",
		"   4b. 合并单元格：跨列合并（如值域横跨多列）与跨行合并（如竖排分组标签跨多行）的范围必须与 PDF 一致；出现多余的空白单元格、合并范围错位等，均判为问题。",
		"   4c. 分组表头：多级表头（如 序号|大类|中类|小类|功能要求|测试结果）的层级归属必须正确，分组标题（如 硬件环境/软件环境/大数据屏）只应覆盖其对应的行组。",
		"   4d. 行数完整性：不应有多余的空白行，也不应缺少 PDF 中存在的行；PDF 中相邻的两行（如 任务来源 行与 客户名称 行）不应被合并到同一行。",
		"   4e. 边框与表线：表格边框、表格内部分隔线是否与 PDF 一致。",
		"5. 图片：图片数量与位置是否一致。",
		"6. 边框与背景：代码块边框、浅灰背景、语法高亮颜色是否保留。",
		"7. 分页差异：Word 预览与 PDF 的分页位置可能不同，属正常现象，不按问题计。",
		"",
		"仅输出一个 JSON 对象，不要输出任何其他文字，格式：",
		'{"match": true|false, "score": 0到1的小数, "issues": [{"severity": "low|medium|high", "description": "问题描述"}], "note": "一句话总评"}',
		"",
		"判定规则：",
		"- score 反映整体保真度（1 = 完全一致）。",
		"- 若 score >= 0.75 且没有 severity=high 的问题，match 必须为 true，否则为 false。",
		"- 表格结构问题（列错位、合并范围错误、缺行/多行、标签与值错位）severity 至少为 medium；导致内容缺失或内容落入错误列的为 high。",
		"- 无问题时 issues 为空数组。"
	].join("\n");
}

/** 从模型输出解析校验判定：剥围栏、截取首个 { 到末个 }、容错字段。 */
function parseVerdict(raw) {
	const fallback = { match: null, score: null, issues: [], note: "no model output" };
	let text = String(raw === undefined || raw === null ? "" : raw).trim();
	if (!text) return fallback;
	const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
	if (fence) text = fence[1].trim();
	const start = text.indexOf("{");
	const end = text.lastIndexOf("}");
	if (start === -1 || end <= start) return { ...fallback, note: "模型输出中未找到 JSON" };
	let data;
	try {
		data = JSON.parse(text.slice(start, end + 1));
	} catch {
		return { ...fallback, note: "模型输出不是合法 JSON" };
	}
	if (!data || typeof data !== "object") return fallback;
	let score = Number(data.score);
	if (!Number.isFinite(score)) score = null;
	else score = Math.min(1, Math.max(0, score));
	let match = data.match;
	if (match !== true && match !== false) match = null;
	const issues = (Array.isArray(data.issues) ? data.issues : []).slice(0, 20).map((i) => {
		const sev = typeof i === "string" ? "medium" : i && typeof i.severity === "string" ? i.severity : "medium";
		const desc = typeof i === "string" ? i : i && typeof i.description === "string" ? i.description : String((i && (i.description ?? i)) ?? "");
		return { severity: sev, description: desc };
	});
	const note = typeof data.note === "string" ? data.note : "";
	if (match === null && score !== null) {
		match = score >= 0.75 && !issues.some((i) => i.severity === "high");
	}
	return { match, score, issues, note };
}

/** 汇总逐页判定 → 报告 markdown（格式与既有报告一致；含 verify→fix 修复循环段）。 */
function buildReport(args) {
	const { pdf, docx, renderer, nPages, nScan, perPage, warnings, overallMatch, avg, fix } = args;
	const L = [];
	L.push("# PDF 转 Word 校验报告", "");
	L.push("- 输入: " + pdf);
	L.push("- 输出: " + docx);
	L.push("- 渲染器: " + (renderer || "unknown"));
	L.push("- 页数: " + nPages + "（扫描页 " + nScan + "）", "");
	L.push("## 总体判定", "");
	L.push("- match: " + (overallMatch === null ? "未校验" : overallMatch ? "通过" : "未通过"));
	if (avg !== null) L.push("- 平均 score: " + round2(avg));
	L.push("- 校验页数: " + perPage.length, "");
	L.push("## 逐页校验", "");
	L.push("| 页 | 模式 | match | score | 主要问题 |");
	L.push("| --- | --- | --- | --- | --- |");
	for (const p of perPage) {
		const m = p.match === true ? "✓" : p.match === false ? "✗" : "–";
		const s = p.score === null || p.score === undefined ? "–" : String(p.score);
		const issueText = p.issues && p.issues.length
			? p.issues.map((i) => (i.severity || "medium") + ":" + (i.description || "")).join("；")
			: p.note || "";
		L.push("| " + p.page + " | " + p.mode + " | " + m + " | " + s + " | " + issueText + " |");
	}
	L.push("");
	L.push("## 修复循环（verify→fix）", "");
	if (!fix || !fix.ran) {
		L.push("（未执行：" + (fix && fix.reason ? fix.reason : "不适用") + "）", "");
	} else if (!fix.rounds || fix.rounds.length === 0) {
		L.push("（循环未产生修复轮次：" + (fix.stopReason || "未知") + "）", "");
	} else {
		L.push("- 停止原因: " + (fix.stopReason || "达到轮次上限"));
		for (const r of fix.rounds) {
			const bits = [
				"轮 " + r.round + "：目标页 " + (r.pages || []).join(","),
				"方案 " + (r.planned || 0) + " 动作",
				"应用 " + (r.applied || 0) + " 动作"
			];
			if (r.rolledBack) bits.push("整表回滚");
			if (Array.isArray(r.reverify) && r.reverify.length) {
				bits.push(
					"复检 " + r.reverify
						.map((v) => "p" + v.page + (v.match === true ? "✓" : v.match === false ? "✗" : "–") + (v.issueCount !== undefined ? "(" + v.issueCount + ")" : ""))
						.join(" ")
				);
			}
			if (r.note) bits.push(r.note);
			L.push("- " + bits.join("；"));
		}
		L.push("");
	}
	L.push("## 转换警告", "");
	if (warnings.length) for (const w of warnings) L.push("- " + w);
	else L.push("（无）");
	L.push("");
	L.push("> 说明：预览渲染器基于 DOCX 声明样式绘制，不模拟 Word 的自动重排与分页；Word 打开时的实际分页可能略有差异，属正常现象。");
	return L.join("\n");
}

function renderResult(value) {
	const L = [];
	L.push("PDF 转 Word 完成：" + value.docx);
	L.push("页数: " + value.n_pages + "（扫描 " + value.n_scan + "），渲染器 " + (value.renderer || "-"));
	if (Array.isArray(value.scan_pages) && value.scan_pages.length) {
		L.push("扫描页: " + value.scan_pages.join(", ") + "（0 基页号，走 OCR）");
	}
	if (value.overall) {
		L.push(
			"校验: " + (value.overall.match ? "通过" : "未通过") +
			"（" + value.overall.verified + " 页，平均 " + round2(value.overall.avg) + "）" +
			(value.report ? "  报告: " + value.report : "")
		);
		if (value.fix) {
			L.push(
				"修复循环: " + value.fix.rounds + " 轮，应用 " + value.fix.applied + " 个动作，" +
				value.fix.stillOpen + " 页仍有未解决的 medium+ 问题"
			);
		}
	} else {
		if (value.report) L.push("报告: " + value.report);
	}
	if (value.warnings && value.warnings.length) {
		L.push("警告:");
		for (const w of value.warnings.slice(0, 10)) L.push("- " + w);
	}
	return [{ type: "text", text: L.join("\n") }];
}

// ---------------------------------------------------------------------------
// apply
// ---------------------------------------------------------------------------
/**
 * @param ctx - Cordis 上下文（含 tools / shell / fs；软依赖 llm / attachments / sandboxPolicy）。
 * @param config - 行配置：verifyProvider、verifyModel、convertTimeoutMs、python。
 */
export function apply(ctx, config) {
	if (!existsSync(CONVERT_PY)) {
		throw new Error("管线缺失：" + CONVERT_PY + " — 请确认包内 pipeline/ 完整（重新安装本包）");
	}
	const llm = ctx.get ? ctx.get("llm") : undefined;
	const attachments = ctx.get ? ctx.get("attachments") : undefined;
	const policy = ctx.get ? ctx.get("sandboxPolicy") : undefined;

	const provider = (config && typeof config.verifyProvider === "string" && config.verifyProvider) || "local";
	const model = (config && typeof config.verifyModel === "string" && config.verifyModel) || "Qwen3.8-27B";
	const timeoutMs = Number(config && config.convertTimeoutMs) > 0 ? Number(config.convertTimeoutMs) : 1800000;

	/** 行配置 autoSetup: false 可禁用首次自动建 .venv。 */
	const autoSetupAllowed = !(config && config.autoSetup === false);
	/** .venv 自动创建的一次性 Promise（防并发重复创建）与耗时记录。 */
	let setupPromise = null;
	let autoSetupMs = 0;

	/** 解释器缺失时（行配置 python / PDF2WORD_PYTHON / 包内 .venv 均未命中），
	 *  首次调用自动在包目录运行 scripts/setup-venv.mjs 建立完整 .venv；
	 *  之后调用直接复用，不再重复安装。 */
	async function ensurePython(exec) {
		const found = resolvePython(config);
		if (found) return found;
		if (!autoSetupAllowed) {
			throw new Error(
				"未找到 Python 解释器：请在插件目录运行 `node scripts/setup-venv.mjs` 建立 .venv，" +
				"或设置环境变量 PDF2WORD_PYTHON / 行配置 python 指向解释器（行配置 autoSetup: false 禁用了自动创建）"
			);
		}
		if (!setupPromise) {
			setupPromise = (async () => {
				const t0 = Date.now();
				let sp;
				if (policy) {
					try {
						const session = sessionOf(exec);
						sp = policy.resolve(session ? { session: session } : {});
					} catch {
						sp = undefined;
					}
				}
				const request = {
					command: '& "' + process.execPath + '" "' + join(PKG_ROOT, "scripts", "setup-venv.mjs") + '"',
					workdir: PKG_ROOT,
					timeoutMs: 1800000,
					stdoutMaxBytes: 8388608
				};
				if (sp) request.sandboxPolicy = sp;
				const result = await ctx.shell.run(ctx.shell.resolve(request));
				if (result.aborted) throw new Error("Python 环境自动创建被中断");
				if (result.timedOut) {
					throw new Error("Python 环境自动创建超时（30 分钟）：请手动在插件目录运行 `node scripts\\setup-venv.mjs`");
				}
				const after = resolvePython(config);
				if (!after) {
					const errText =
						(result.stderr && result.stderr.text) || (result.stdout && result.stdout.text) || "";
					throw new Error(
						"Python 环境自动创建后仍无可用解释器（本机装有 Python 3.10–3.12 吗？）：" +
						"请安装 Python 3.12 后手动运行 `node scripts\\setup-venv.mjs`" +
						(errText ? "；输出尾部：" + tail(errText, 300) : "")
					);
				}
				autoSetupMs = Date.now() - t0;
				return after;
			})().catch((error) => {
				setupPromise = null;
				throw error;
			});
		}
		return setupPromise;
	}

	/** 通过 shell 服务启动 Python 管线（按会话解析沙箱策略）。 */
	async function runConvert(args, exec) {
		const python = await ensurePython(exec);
		const cwd = sessionCwd(exec) || process.cwd();
		let sp;
		if (policy) {
			try {
				const session = sessionOf(exec);
				sp = policy.resolve(session ? { session: session } : {});
			} catch {
				sp = undefined;
			}
		}
		const parts = [
			'"' + python + '"',
			'"' + CONVERT_PY + '"',
			"--pdf", '"' + args.pdf + '"',
			"--out", '"' + args.out + '"',
			"--mode", args.mode,
			"--assets", '"' + args.assets + '"'
		];
		const request = {
			command: "& " + parts.join(" "),
			workdir: cwd,
			timeoutMs: timeoutMs,
			stdoutMaxBytes: 16777216,
			...(exec && exec.signal ? { signal: exec.signal } : {})
		};
		if (sp) request.sandboxPolicy = sp;
		const result = await ctx.shell.run(ctx.shell.resolve(request));
		if (result.aborted) throw new Error("转换已中断");
		if (result.timedOut) throw new Error("转换超时（" + timeoutMs + " ms）");
		const out = (result.stdout && result.stdout.text) || "";
		const parsed = parsePipelineJson(out);
		if (!parsed) {
			const errText = (result.stderr && result.stderr.text) || out;
			throw new Error("管线未返回 PDF2WORD_JSON 结果" + (errText ? "；输出尾部：" + tail(errText, 500) : ""));
		}
		if (parsed.ok !== true) {
			const errText = (result.stderr && result.stderr.text) || "";
			throw new Error("管线转换失败：" + (parsed.error || "unknown") + (errText ? "；stderr：" + tail(errText, 300) : ""));
		}
		return parsed;
	}

	/** 运行包内修复脚本（fixdocx.py / preview.py），返回 stdout 文本。
	 *  与 runConvert 共用解释器解析与沙箱策略；独立超时 300s。 */
	async function runFixTool(scriptPath, argString, exec) {
		const python = await ensurePython(exec);
		const cwd = sessionCwd(exec) || process.cwd();
		let sp;
		if (policy) {
			try {
				const session = sessionOf(exec);
				sp = policy.resolve(session ? { session: session } : {});
			} catch {
				sp = undefined;
			}
		}
		const request = {
			command: '& "' + python + '" "' + scriptPath + '" ' + argString,
			workdir: cwd,
			timeoutMs: 300000,
			stdoutMaxBytes: 16777216,
			...(exec && exec.signal ? { signal: exec.signal } : {})
		};
		if (sp) request.sandboxPolicy = sp;
		const result = await ctx.shell.run(ctx.shell.resolve(request));
		if (result.aborted) throw new Error("修复脚本被中断");
		if (result.timedOut) throw new Error("修复脚本超时（300 秒）");
		return (result.stdout && result.stdout.text) || "";
	}

	/** 读取图片字节（经 fs 服务）。 */
	async function readImageBytes(absPath, maxBytes, signal) {
		const t = await ctx.fs.resolve(absPath);
		const st = await ctx.fs.stat(t, signal);
		if (!st) throw new Error("文件不存在：" + absPath);
		return await ctx.fs.readBytes(t, signal, maxBytes);
	}

	/** 调用 llm 服务流式生成并收集文本。maxTokens 缺省 2048（校验判定够用）；
	 *  修复方案含整格文本，需放大（8192）。 */
	async function llmText(prompt, imageRefs, signal, maxTokens) {
		const message = createMessage({
			role: "user",
			content: [
				{ type: "text", text: prompt },
				...imageRefs.map((ref) => ({ type: "image", attachment: ref }))
			],
			source: { kind: "user" }
		});
		let text = "";
		let finishReason = null;
		for await (const chunk of llm.stream({
			provider: provider,
			model: model,
			messages: [message],
			temperature: 0.1,
			maxTokens: Number(maxTokens) > 0 ? Math.floor(Number(maxTokens)) : 2048,
			...(signal ? { signal: signal } : {})
		})) {
			if (chunk.type === "text-delta") text += chunk.text;
			else if (chunk.type === "finish") finishReason = chunk.reason;
		}
		if (
			finishReason &&
			finishReason.kind !== "stop" &&
			finishReason.kind !== "max-tokens" &&
			finishReason.kind !== "tool-calls"
		) {
			throw new Error(
				"LLM 校验异常结束（" + finishReason.kind + "）：" +
				((finishReason.failure && finishReason.failure.message) || "无详情")
			);
		}
		if (!text.trim()) throw new Error("LLM 校验返回空内容");
		return text;
	}

	const PAGE_ITEM_SCHEMA = {
		type: "object",
		additionalProperties: false,
		properties: {
			page: { type: "number" },
			mode: { type: "string" },
			match: { type: "boolean" },
			score: { type: "number" },
			issues: {
				type: "array",
				items: {
					type: "object",
					additionalProperties: false,
					properties: {
						severity: { type: "string" },
						description: { type: "string" }
					}
				}
			},
			note: { type: "string" }
		}
	};

	ctx.tools.register(defineTool({
		name: "pdf_to_word",
		description:
			"将 PDF 转换为 Word(.docx)，尽量保持字体大小、段落样式、表格、图片等格式；" +
			"兼容带文字层的数字版 PDF 与扫描件 PDF（自动按页检测，扫描件走 PaddleOCR）。" +
			"转换后可用多模态大模型逐页比对 PDF 与 Word 渲染图，输出样式一致性校验报告。" +
			"校验发现 medium+ 表格结构问题后自动进入 verify→fix 闭环（LLM 出白名单修复方案，" +
			"确定性脚本执行并带不变式校验与整表回滚，重渲染后仅复检受影响页，直至通过或达轮次上限）。" +
			"参数: pdfPath(必填, PDF 路径), mode(auto|digital|scan, 默认 auto), " +
			"verify(是否多模态校验, 默认 true), maxVerifyPages(默认 8), " +
			"maxFixRounds(修复轮次上限, 0=禁用, 默认 2, 上限 3), outPath(可选输出 .docx 路径)。",
		parameters: {
			pdfPath: {
				type: "string",
				required: true,
				description: "输入 PDF 文件路径（绝对路径或相对路径）。"
			},
			mode: {
				type: "string",
				enum: ["auto", "digital", "scan"],
				description: "转换模式；auto=按页自动检测文字层"
			},
			verify: {
				type: "boolean",
				description: "是否执行多模态大模型样式校验，默认 true"
			},
			maxVerifyPages: {
				type: "number",
				description: "最多校验的页数，默认 8"
			},
			maxFixRounds: {
				type: "number",
				description: "校验发现 medium+ 表格结构问题后的自动修复轮次上限；0=禁用，默认 2，上限 3"
			},
			outPath: {
				type: "string",
				description: "输出 .docx 路径；默认与输入同目录同名"
			}
		},
		output: {
			schema: {
				type: "object",
				additionalProperties: false,
				properties: {
					ok: { type: "boolean" },
					docx: { type: "string" },
					report: { type: "string" },
					renderer: { type: "string" },
					n_pages: { type: "number" },
					n_scan: { type: "number" },
					scan_pages: { type: "array", items: { type: "number" } },
					overall: {
						type: "object",
						additionalProperties: false,
						properties: {
							match: { type: "boolean" },
							avg: { type: "number" },
							verified: { type: "number" }
						}
					},
					pages: { type: "array", items: PAGE_ITEM_SCHEMA },
					fix: {
						type: "object",
						additionalProperties: false,
						properties: {
							rounds: { type: "number" },
							applied: { type: "number" },
							stillOpen: { type: "number" }
						}
					},
					warnings: { type: "array", items: { type: "string" } }
				}
			},
			render: (_args, value) => renderResult(value)
		},
		async execute(args, exec) {
			const signal = exec && exec.signal ? exec.signal : undefined;
			const cwd = sessionCwd(exec) || process.cwd();

			// 1. 路径解析
			const pdfInput = String(args.pdfPath || "").trim();
			if (!pdfInput) throw new Error("pdfPath 不能为空");
			const pdf = normalize(isAbsolute(pdfInput) ? pdfInput : join(cwd, pdfInput));
			const pdfTarget = await ctx.fs.resolve(pdf, { cwd: cwd });
			const pdfStat = await ctx.fs.stat(pdfTarget, signal);
			if (!pdfStat) throw new Error("PDF 不存在：" + pdf);
			if (pdfStat.type !== "file") throw new Error("不是文件：" + pdf);

			let outPath = args.outPath !== undefined && args.outPath !== null ? String(args.outPath).trim() : "";
			if (outPath && !isAbsolute(outPath)) outPath = join(cwd, outPath);
			outPath = outPath ? normalize(outPath) : pdf.replace(/\.pdf$/i, "") + ".docx";
			const assetsDir = join(dirname(pdf), "pdf2w_assets");
			const mode = args.mode || "auto";
			const wantVerify = args.verify !== false;
			const maxPages = Number(args.maxVerifyPages) > 0 ? Math.floor(Number(args.maxVerifyPages)) : 8;

			// 2. 转换（首次调用可能触发 .venv 自动创建）
			const result = await runConvert({ pdf: pdf, out: outPath, mode: mode, assets: assetsDir }, exec);
			const warnings = (Array.isArray(result.warnings) ? result.warnings : []).map(String);
			if (autoSetupMs > 0) {
				warnings.push(
					"首次调用已自动创建 Python 环境 .venv（耗时约 " + Math.ceil(autoSetupMs / 1000) + " 秒）；" +
					"注意 git 安装的包目录位于 pnpm 存储区内，重装/更新后 .venv 可能需重建（也可用外部 venv + PDF2WORD_PYTHON 指向更稳定的位置）"
				);
			}
			const pagesRaw = Array.isArray(result.pages) ? result.pages : [];
			const scanPages = Array.isArray(result.scan_pages)
				? result.scan_pages.map((p) => Number(p)).filter((p) => Number.isFinite(p) && p >= 0)
				: [];
			if (scanPages.length > 0) {
				warnings.push(
					"检测到扫描件页面 " + scanPages.length + " 页（0 基页号：" + scanPages.join(", ") + "），" +
					"已走 OCR 识别；OCR 文本无原始文字层，字体/字号为版面估计值"
				);
			}

			// 3. 校验（可选）
			let perPage = [];
			let overallMatch = null;
			let avg = null;
			let reportPath = null;
			const maxFixRounds = args.maxFixRounds === undefined || args.maxFixRounds === null
				? 2
				: Math.min(3, Math.max(0, Math.floor(Number(args.maxFixRounds) || 0)));
			/** verify→fix 闭环状态（fixRounds 记录每轮，写入报告与 warnings）。 */
			const fix = { ran: false, reason: "", rounds: [], applied: 0, stopReason: "" };
			if (wantVerify) {
				if (!llm || !attachments) {
					warnings.push("未启用校验：宿主缺少 llm 或 attachments 服务");
				} else {
					const limits = attachments.imageLimits || {};
					const maxImageBytes = Number.isFinite(limits.maxImageBytes) ? limits.maxImageBytes : 10485760;
					const selected = pickPages(pagesRaw, maxPages);
					for (const p of selected) {
						const entry = pagesRaw.find((x) => Number(x.page) === p);
						const pMode = entry && entry.mode ? String(entry.mode) : mode;
						const pdfImg = join(assetsDir, "pdf_p" + p + ".jpg");
						const wordImg = join(assetsDir, "word_p" + p + ".jpg");
						let verdict;
						try {
							const b1 = await readImageBytes(pdfImg, maxImageBytes, signal);
							const b2 = await readImageBytes(wordImg, maxImageBytes, signal);
							const ref1 = await attachments.saveImage({ data: b1, mediaType: "image/jpeg", name: "pdf_p" + p + ".jpg" });
							const ref2 = await attachments.saveImage({ data: b2, mediaType: "image/jpeg", name: "word_p" + p + ".jpg" });
							const text = await llmText(verifyPrompt(p + 1), [ref1, ref2], signal);
							verdict = parseVerdict(text);
						} catch (error) {
							verdict = {
								match: false,
								score: null,
								issues: [{ severity: "high", description: "校验失败：" + (error && error.message ? error.message : String(error)) }],
								note: ""
							};
						}
						perPage.push({
							page: p + 1,
							mode: pMode,
							...(verdict.match !== null ? { match: verdict.match } : {}),
							...(verdict.score !== null ? { score: verdict.score } : {}),
							issues: verdict.issues,
							note: verdict.note
						});
					}
					if (perPage.length > 0) {
						overallMatch = perPage.every((p) => p.match === true);
						const scores = perPage.map((p) => p.score).filter((s) => s !== null && s !== undefined);
						avg = scores.length > 0 ? scores.reduce((a, b) => a + b, 0) / scores.length : null;
					}

					// 4. verify→fix 闭环（仅当存在 medium+ 表格结构问题时）
					if (maxFixRounds > 0 && perPage.some(pageHasFixableIssues)) {
						if (!existsSync(FIXDOCX_PY)) {
							warnings.push("修复循环未执行：管线缺失 " + FIXDOCX_PY + "（重新安装本包）");
						} else {
							fix.ran = true;
							let round = 0;
							while (round < maxFixRounds) {
								const bad = perPage.filter(pageHasFixableIssues);
								if (bad.length === 0) {
									fix.stopReason = "无 medium+ 问题残留";
									break;
								}
								round += 1;
								// 选最差的 ≤3 页：high 优先 → 分数低 → 页序
								const worst = [...bad].sort((a, b) => {
									const ha = a.issues.some((i) => i.severity === "high") ? 0 : 1;
									const hb = b.issues.some((i) => i.severity === "high") ? 0 : 1;
									if (ha !== hb) return ha - hb;
									const sa = a.score === null || a.score === undefined ? 2 : a.score;
									const sb = b.score === null || b.score === undefined ? 2 : b.score;
									if (sa !== sb) return sa - sb;
									return a.page - b.page;
								}).slice(0, 3);
								const rec = { round: round, pages: worst.map((p) => p.page), planned: 0, applied: 0, rolledBack: false, reverify: [], note: "" };
								fix.rounds.push(rec);
								try {
									// 4.1 导出表格结构
									const dumpOut = await runFixTool(FIXDOCX_PY, "dump \"" + outPath + "\"", exec);
									let dumpText = dumpOut.trim();
									JSON.parse(dumpText); // 合法性
									if (dumpText.length > 60000) dumpText = dumpText.slice(0, 60000) + " …(截断)";
									// 4.2 问题清单（上轮校验发现）
									const pageIssues = worst.flatMap((p) =>
										p.issues.map((i) => "- 第 " + p.page + " 页 [" + i.severity + "] " + i.description)
									);
									// 4.3 附图：每页 PDF 原图 + 最新 Word 预览（各一对）
									const imageRefs = [];
									for (const p of worst) {
										const idx = p.page - 1;
										const b1 = await readImageBytes(join(assetsDir, "pdf_p" + idx + ".jpg"), maxImageBytes, signal);
										const b2 = await readImageBytes(join(assetsDir, "word_p" + idx + ".jpg"), maxImageBytes, signal);
										imageRefs.push(await attachments.saveImage({ data: b1, mediaType: "image/jpeg", name: "fix_r" + round + "_pdf_p" + idx + ".jpg" }));
										imageRefs.push(await attachments.saveImage({ data: b2, mediaType: "image/jpeg", name: "fix_r" + round + "_word_p" + idx + ".jpg" }));
									}
									// 4.4 LLM 出方案（maxTokens 放大以容纳整格文本）
									const planRaw = await llmText(
										planFixPrompt({ dumpText, pageIssues, worstPages: worst.map((p) => p.page) }),
										imageRefs,
										signal,
										8192
									);
									const actions = parseFixPlan(planRaw, warnings);
									rec.planned = actions.length;
									if (actions.length === 0) {
										rec.note = "方案未给出可修复动作";
										fix.stopReason = "模型判定无可修复的表格结构问题";
										break;
									}
									const planPath = join(assetsDir, "fix_plan_r" + round + ".json");
									await ctx.fs.writeText(
										await ctx.fs.resolve(planPath, { cwd: cwd }),
										JSON.stringify({ actions: actions, notes: "verify→fix round " + round }, null, 1),
										undefined,
										signal
									);
									// 4.5 确定性应用（不变式门禁 + 整表回滚；成功才写 .pre_fix.docx 备份）
									const applyOut = await runFixTool(
										FIXDOCX_PY,
										"apply \"" + outPath + "\" --plan \"" + planPath + "\" --in-place",
										exec
									);
									const applyRes = parseFixJson(applyOut);
									if (!applyRes) throw new Error("修复脚本未返回 FIXDOCX_JSON 结果");
									rec.applied = Number(applyRes.applied) || 0;
									rec.rolledBack = applyRes.rolled_back === true;
									if (Array.isArray(applyRes.failed) && applyRes.failed.length > 0) {
										rec.note = applyRes.failed.slice(0, 3)
											.map((f) => f.op + "(" + (f.reason || "") + ")")
											.join("；") + " 未通过校验";
									}
									if (rec.rolledBack) {
										fix.stopReason = "不变式校验失败，已整表回滚" + (applyRes.invariants ? "：" + applyRes.invariants : "");
										rec.note = (rec.note ? rec.note + "；" : "") + fix.stopReason;
										break;
									}
									if (rec.applied === 0) {
										fix.stopReason = "无动作成功应用" + (applyRes.error ? "：" + applyRes.error : "");
										break;
									}
									fix.applied += rec.applied;
									// 4.6 重渲染 Word 预览
									await runFixTool(PREVIEW_PY, "\"" + outPath + "\" \"" + assetsDir + "\" --dpi 140", exec);
									// 4.7 仅复检受影响页，更新 perPage 与总体判定
									for (const p of worst) {
										const idx = p.page - 1;
										let verdict;
										try {
											const b1 = await readImageBytes(join(assetsDir, "pdf_p" + idx + ".jpg"), maxImageBytes, signal);
											const b2 = await readImageBytes(join(assetsDir, "word_p" + idx + ".jpg"), maxImageBytes, signal);
											const ref1 = await attachments.saveImage({ data: b1, mediaType: "image/jpeg", name: "reverify_r" + round + "_pdf_p" + idx + ".jpg" });
											const ref2 = await attachments.saveImage({ data: b2, mediaType: "image/jpeg", name: "reverify_r" + round + "_word_p" + idx + ".jpg" });
											const text = await llmText(verifyPrompt(idx + 1), [ref1, ref2], signal);
											verdict = parseVerdict(text);
										} catch (error) {
											verdict = {
												match: false,
												score: null,
												issues: [{ severity: "high", description: "校验失败：" + (error && error.message ? error.message : String(error)) }],
												note: ""
											};
										}
										const entry = perPage.find((x) => x.page === p.page);
										if (entry) {
											if (verdict.match === true || verdict.match === false) entry.match = verdict.match;
											if (verdict.score !== null && verdict.score !== undefined) entry.score = verdict.score;
											entry.issues = verdict.issues;
											entry.note = verdict.note;
										}
										rec.reverify.push({ page: p.page, match: verdict.match, issueCount: verdict.issues.length });
									}
									overallMatch = perPage.every((p) => p.match === true);
									const scores = perPage.map((p) => p.score).filter((s) => s !== null && s !== undefined);
									avg = scores.length > 0 ? scores.reduce((a, b) => a + b, 0) / scores.length : null;
								} catch (error) {
									rec.note = (rec.note ? rec.note + "；" : "") + "轮次异常：" + (error && error.message ? error.message : String(error));
									fix.stopReason = rec.note;
									warnings.push("修复轮 " + round + " 异常，停止循环：" + (error && error.message ? error.message : String(error)));
									break;
								}
							}
							if (!fix.stopReason) fix.stopReason = "达到轮次上限 " + maxFixRounds;
							if (fix.applied > 0) {
								warnings.push("verify→fix 闭环已应用 " + fix.applied + " 个修复动作（" + fix.rounds.length + " 轮）；输出文件旁保留 .pre_fix.docx 备份，如需回退请手动恢复");
							}
						}
					} else if (maxFixRounds === 0 && perPage.some(pageHasFixableIssues)) {
						fix.reason = "已按 maxFixRounds=0 禁用";
					}

					// 5. 写报告
					const reportContent = buildReport({
						pdf: pdf,
						docx: outPath,
						renderer: result.renderer,
						nPages: Number(result.n_pages) || 0,
						nScan: Number(result.n_scan) || 0,
						perPage: perPage,
						warnings: warnings,
						overallMatch: overallMatch,
						avg: avg,
						fix: fix
					});
					reportPath = join(assetsDir, "verify_report.md");
					try {
						const rt = await ctx.fs.resolve(reportPath, { cwd: cwd });
						await ctx.fs.writeText(rt, reportContent, undefined, signal);
					} catch (error) {
						reportPath = null;
						warnings.push("报告写入失败：" + (error && error.message ? error.message : String(error)));
					}
				}
			}

			return {
				ok: true,
				docx: outPath,
				...(reportPath ? { report: reportPath } : {}),
				renderer: result.renderer ? String(result.renderer) : "",
				n_pages: Number(result.n_pages) || 0,
				n_scan: Number(result.n_scan) || 0,
				scan_pages: scanPages,
				...(overallMatch !== null
					? { overall: { match: overallMatch, avg: avg === null ? 0 : round2(avg), verified: perPage.length } }
					: {}),
				pages: perPage,
				...(wantVerify
					? { fix: { rounds: fix.rounds.length, applied: fix.applied, stillOpen: perPage.filter(pageHasFixableIssues).length } }
					: {}),
				warnings: warnings
			};
		}
	}));
}
