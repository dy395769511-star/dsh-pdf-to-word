# dsh-pdf-to-word

DeepSeek Harness（dsh）工程化插件：把 PDF 转成 Word（.docx），尽量保真字体大小、段落样式、表格、图片、边框/背景；扫描件自动走 PaddleOCR；转换后可用多模态大模型逐页比对并输出样式一致性校验报告。

注册一个模型工具 **`pdf_to_word`**，行为与 dsh 会话内的动态插件版完全一致。

## 目录结构

```
pdf-to-word/
├── package.json            # 包元数据（type: module, dsh.bundle.patch）
├── cordis.patch.yml        # 插件行（id/name/config），add 时并入 profile bundles
├── lib/
│   └── index.js            # 静态 Cordis 插件入口（注册 pdf_to_word 工具）
├── pipeline/
│   ├── convert.py          # 主转换（数字版保真：字号/表格/图片/边框/背景）
│   ├── fixdocx.py          # 表格结构 dump/apply（verify→fix 闭环的确定性执行器）
│   ├── preview.py          # DOCX 预览渲染（供校验比对的 word_pN.jpg；修复循环重渲染）
│   ├── scanocr.py          # 扫描页 OCR（PaddleOCR 3.x / PP-OCRv6）
│   └── requirements.txt    # Python 依赖
├── scripts/
│   └── setup-venv.mjs      # 建立 .venv 并安装依赖（node scripts/setup-venv.mjs）
└── .venv/                  # Python 虚拟环境（本机 junction 或自建，不入库）
```

## 安装

### 1. 安装到 profile

**推荐：从 GitHub 直接安装**（仓库：[dy395769511-star/dsh-pdf-to-word](https://github.com/dy395769511-star/dsh-pdf-to-word)，公共仓库，无需登录）：

```bat
dsh plugin --profile web add github:dy395769511-star/dsh-pdf-to-word
```

或显式 git URL（等价）：

```bat
dsh plugin --profile web add dsh-pdf-to-word@https://github.com/dy395769511-star/dsh-pdf-to-word.git
```

- 需要本机装有 `pnpm` 与 `git`（`dsh plugin add` 是 pnpm 转发器，git 依赖经 git 拉取）。
- 默认取 `main` 分支 HEAD；可加 `#<tag 或 commit>` 固定版本，如 `...dsh-pdf-to-word.git#v1.0.0`。
- peer 依赖 `@deepseek-ai/dsh-tools` / `@deepseek-ai/dsh-llm`（运行时实际由 DSH 宿主提供）安装时由 pnpm 从 npm registry 自动补齐，无需额外操作。
- `add` 完成后 launcher 自动把 `cordis.patch.yml` 中的行并入 profile 的 `dsh.profile.bundles`。

安装后包位于（`%DSH_HOME%` 默认 `%USERPROFILE%\.dsh`）：

```
%DSH_HOME%\profiles\web\node_modules\dsh-pdf-to-word\
```

其他安装方式（与 GitHub 安装二选一）：

- `link:<绝对路径>`：本地开发（软链，改动即时生效，需重启 dsh 加载），如
  `dsh plugin --profile web add link:E:\2026\dsh\plugins\pdf-to-word`。
- tarball：`pnpm pack` 出真实文件后 `dsh plugin --profile web add <tarball>`（任意位置）。

### 2. Python 环境（默认自动创建）

需要 Python 3.10–3.12（管线基于 3.12 构建）。**通常无需手动准备**：首次调用 `pdf_to_word` 时若未找到可用解释器，插件会自动在包目录运行 `setup-venv.mjs` 创建完整 `.venv`（含 Paddle OCR，约 500MB，耗时视网络而定，本机 pip 缓存热时约 3 分钟），完成后继续本次转换，后续调用直接复用。

也可以提前在**包目录**（本地 checkout 目录或上一步的安装目录）准备（更快看到进度 / 只需数字版时用 `--core`）：

```bat
node scripts\setup-venv.mjs          :: 完整安装（含 Paddle OCR，约 500MB）
node scripts\setup-venv.mjs --core   :: 仅数字版转换（不含扫描 OCR）
```

生成 `<包目录>/.venv`。也可用现成解释器：设环境变量 `PDF2WORD_PYTHON` 或在插件行配置 `python` 指向它（见“配置”）。

> GitHub 安装的包目录位于 pnpm 存储区内，其中的 `.venv` 在 profile 重装/包更新时可能被清理。
> 若追求稳定：把 venv 建在固定外部位置，再用 `PDF2WORD_PYTHON`（或行配置 `python`）指向它，例如
> `C:\tools\python312\python.exe -m venv C:\tools\pdf2word-venv` 后
> `C:\tools\pdf2word-venv\Scripts\pip install -r <包目录>\pipeline\requirements.txt`，
> 设 `PDF2WORD_PYTHON=C:\tools\pdf2word-venv\Scripts\python.exe`。
>
> 本机开发捷径：若已有 venv，可直接建 junction，例如
> `cmd /c mklink /J <包目录>\.venv E:\2026\dsh\.dsh-pdf2word\.venv`。
> 扫描模式模型缓存同理可 junction 到 `.paddlex-cache/`（或用 `PADDLEX_HOME` 指向）。

### 3. 重启 dsh 生效

插件行在 profile 启动时挂载：重启 `dsh web`（或对应 profile 进程）后新会话即可用 `pdf_to_word` 工具。

## 前置条件（校验功能）

- dsh 设置「模型」中已配置一个**多模态** provider（本部署默认 `local` / `Qwen3.8-27B`，凭据走 dsh 设置与 `.credentials.yaml`，不进入本包）。
- 若只装 `--core` 或无 LLM 服务，工具仍可用（转换正常，校验自动跳过并给警告）。

## 工具用法

```
pdf_to_word(
  pdfPath,          # 必填，PDF 路径
  mode?,            # auto(默认) | digital | scan
  verify?,          # 默认 true，是否 LLM 校验
  maxVerifyPages?,  # 默认 8
  maxFixRounds?,    # 默认 2，verify→fix 修复轮次上限（0=禁用，上限 3）
  outPath?          # 输出 .docx 路径，默认与 PDF 同目录同名
)
```

返回：`docx` 输出路径、页数/扫描页数/渲染器、`scan_pages`（auto 模式逐页判定的扫描页 0 基页号列表，非空即含扫描件）、总体判定（match + 平均分）、逐页判定、`fix`（`{rounds, applied, stillOpen}`，修复循环统计）、警告；
校验报告写到 `<pdf 所在目录>/pdf2w_assets/verify_report.md`，修复方案落盘 `pdf2w_assets/fix_plan_rN.json`。

### 判定规则

- 选页：必含首页/末页，其余按 `表格数*2 + 图片数` 降序补足到 `maxVerifyPages`。
- 逐页：模型输出 JSON `{match, score, issues[], note}`；`score ≥ 0.75 且无 high 问题` ⇒ 该页通过。
- 总体：**所有已校验页均通过**才算通过（平均分仅作展示）。
- Word 重排导致的分页差异属正常，不计问题。

## verify→fix 闭环（v1.1）

校验发现 medium+ 表格结构问题时自动进入修复循环（需要 llm + attachments 服务；`maxFixRounds=0` 可禁用）：

1. **选页**：取问题最重的 ≤3 页（high 优先 → 分数低 → 页序）。
2. **dump**：`fixdocx.py dump <docx>` 导出全部表格结构（grid 列宽、每行单元格的文本/跨列/纵向合并），作为模型的精确结构视图。
3. **plan**：多模态 LLM（PDF 原图 + Word 预览图 + 结构转储 + 问题清单）输出严格 JSON 动作方案，动作限于 7 类白名单表格操作：
   `setGridSpan` / `setVMerge` / `setCellText` / `insertCell` / `removeCell` / `cloneRow` / `removeRow`；
   硬约束：终态每行跨列和 == grid 列数、vMerge continue 上方必有 restart/continue、只动问题相关单元格、文本必须取自 PDF。
4. **apply**：`fixdocx.py apply <docx> --plan <json> --in-place` 确定性执行——逐操作在深拷贝草稿上校验，全部应用后做全局不变式门禁（行跨列和、vMerge 良构、文件可重开）；**任一失败则整表回滚、不落盘**；成功则写 `.pre_fix.docx` 备份。
5. **re-verify**：`preview.py` 重渲染 Word 预览，**仅复检本轮受影响页**，更新逐页/总体判定。
6. **循环**：至无 medium+ 问题残留、模型无可修复方案、无动作成功应用、回滚或达到 `maxFixRounds`（默认 2，上限 3）。

修复循环段落写入 `verify_report.md`；每轮方案留存 `pdf2w_assets/fix_plan_rN.json` 供审计/重放。

## 配置（cordis.patch.yml 行 `config:`）

| 键 | 默认 | 说明 |
| --- | --- | --- |
| `verifyProvider` | `local` | 校验 LLM provider（须在 dsh「模型」设置中注册） |
| `verifyModel` | `Qwen3.8-27B` | 校验 LLM 模型 |
| `convertTimeoutMs` | `1800000` | 单次转换墙钟上限（ms） |
| `python` | — | Python 解释器绝对路径（优先级最高） |
| `autoSetup` | `true` | 置 `false` 禁用首次调用的 `.venv` 自动创建 |

Python 解释器解析顺序：行配置 `python` → 环境变量 `PDF2WORD_PYTHON` → `<包目录>/.venv`；均未命中时自动创建（除非 `autoSetup: false`）。

## 工作原理

- 插件通过 **shell 服务** 按会话解析沙箱策略后启动 `<python> pipeline/convert.py --pdf … --out … --mode … --assets <pdf目录>/pdf2w_assets`；
  结果取 stdout 最后一行 `PDF2WORD_JSON <json>`。
- 校验经 **llm 服务**（`ctx.llm.stream`）+ **attachments 服务**（图片落库）：
  逐页读取 `pdf_pN.jpg` 与 `word_pN.jpg`（0 基），两图并送模型比对。
- 报告经 **fs 服务** 写盘。所有服务缺失时优雅降级（转换仍可用）。

## 卸载

```bat
dsh plugin --profile web remove dsh-pdf-to-word
```

launcher 会同步清理 profile 中的 bundle 行。

## 常见问题

- **“Python 环境自动创建后仍无可用解释器”**：本机缺 Python 3.10–3.12。装好 Python 3.12 后重试（自动重建），或手动 `node scripts/setup-venv.mjs` / 设 `PDF2WORD_PYTHON`。
- **“LLM 校验异常结束 / provider 未注册”**：检查 dsh 设置「模型」里 provider id 与行配置 `verifyProvider` 一致。
- **扫描模式首次运行慢**：PaddleOCR 模型按 `PADDLE_PDX_CACHE_HOME` 缓存（默认 `~/.paddlex`，可用包内 `.paddlex-cache/` 或 `PADDLEX_HOME` 覆盖）。
- **扫描 OCR 结果不稳定（同一 PDF 两次转换表格结构不同）**：多线程 CPU 推理的浮点归约顺序不确定。管线默认单线程（`PDF2WORD_OCR_THREADS=1`）保证可复现；如需提速可设 `PDF2WORD_OCR_THREADS` 为核数，代价是结果可能轻微抖动（个别 rowspan/空单元格）。
- **单字符单元格丢失（如表头“序号”列的“1”）或个别密集表格页结构退化**：OCR 输入 JPEG 质量影响 SLANeXt 的结构识别。默认 q85（过低会糊掉单字符格）；可用 `PDF2WORD_OCR_JPEG_QUALITY` 调全局质量，`PDF2WORD_OCR_JPEG_QUALITY_PAGES="6:75"` 按 0 基页号对个别退化页覆盖（实测某密集表格页在 q85 下列结构错乱，回退 q75 后恢复）。
- **sandbox 拦截**：插件走宿主 shell 服务，按其会话沙箱策略执行；只读目录/网络受限环境请在 dsh 设置中放行对应范围。
