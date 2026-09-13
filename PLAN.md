# WordToHtml 编辑器大修工作计划（W3–W5）

> **这份文件是写给「下一个 session」的。** 读完它就应该能直接开工，不必回头翻对话记录。
>
> 来源需求：`issues/20260913-2.md` 末尾「文本编辑器大修（先不动，留待下一个大版本）」那 8 组功能。
> 该清单 2026-09-13 被用户解冻为下一阶段工作，并定下「先小后大」的波次：**W1** 快捷键 + 特殊空格 →
> **W2** 样式模板绑定 + 双页并排 → **W3** 查找替换 + 导航窗格 → **W4** 表格编辑器 → **W5** 节编辑框架。
>
> **维护约定**：每完成一波，把该波标题改成「（已完成 YYYY-MM-DD）」，并在最后一节补一行结论。
> 涉及改数据格式的波次（W4、W5）**必须先出模型设计给用户过目，用户点头才写实现** —— 这是用户定下的闸门。

---

## 0. 工作方式（硬性，先读这段）

**调度方式**：每波**先派实现代理(独立的 headless CLI 进程) → 回来再派独立验收代理**，验收过了才 `git commit`，然后才进下一波。

> 原因：当前subagent几乎100%会崩溃，见第 9 节第 1 条：宿主 TUI 会崩在「大块输出渲进 TUI」上，而 subagent 的报告走的正是 `<task-notification>` → 渲染进主 session 这条路径（三次崩溃全部落在报告落地那一刻）。
> headless 进程没有 TUI，把 stdout 重定向到文件之后，进主 session 的只有 shell 工具那几行摘要。
>
> ```
> qwen -p "<任务书>" --model <与主 session 同一个模型> -y > .qwen/tmp/wX-impl.log 2>&1
> ```
>
> - 任务书里固定要求：**结论写进 `.qwen/tmp/wX-impl-result.md`（≤ 40 行）**，细节与长日志留在
>   `wX-impl.log`；主 session 只读那个结果文件，需要细节再 grep 日志。
> - **要 `-y`（YOLO），不要用 `--approval-mode auto-edit`**：`auto-edit` 只自动批准**文件写入**，
>   **不批准 shell** —— 非交互 session 里 `run_shell_command` 会被直接拒掉（日志里只有一句
>   `Warning: Tool "run_shell_command" requires user approval but cannot execute in non-interactive mode`）。
>   W3 因此交出过一次「代码写完了、命令一条没跑」的成果，类型错误还留在里面（主 session 补跑才发现）。
>   纯写文件的活儿 `auto-edit` 够用；只要任务里含「跑 npm / 起浏览器 / 跑探针」，一律 `-y`。
> - **验收代理同样要 `-y`**（它要跑 `verify:p2`/`verify:p3`、常要自写 playwright 探针），但任务书里必须
>   限定「只许在 `.qwen/tmp/` 下新建文件，仓库内任何文件都不许改」，并且**验收一跑完主 session 立刻
>   `git status --short` 复核**：只要有一个仓库内文件被动过，这次验收作废。
>   若这一轮验收既不用跑命令也不用写脚本，`--approval-mode plan` 更省心 —— 实测它是硬闸门：写操作被拦在
>   执行前，当轮工具表里连 `write_file` / `Edit` 都没有，而且不会挂住（写完「待批准方案」就退出）。
> - 可选护栏：`--max-session-turns N`（超限退出码 53）、`--max-wall-time`、`--max-tool-calls`（预算类退出码 55）。
> - 子进程启动会打一条「MCP server(s) failed to start」（那三个 `qwen-mm-plugins-*` 起不来），
>   **属正常**，内置工具照常可用，不要去修它。
> - 日志是 UTF-8：**用 `read_file` 看，不要用 PowerShell 的 `Get-Content`**（默认 GBK，会花屏）。
> - 可以后台跑（`is_background`），但 stdout 必须重定向到文件 —— 完成通知有可能带上输出尾部。
> - **想彻底绕开「后台任务完成 → 自动续跑」那条不稳的路径**，就用「完全脱离」的起法：`Start-Process -WindowStyle Hidden`
>   （**不要用 `start /b`**，它会被 shell 工具的超时连带杀掉），代价是失败 / 挂住没有信号，只能轮询。
>   完整配方与两个实测坑见 6.8。

**为什么并行度低**：这个仓库的功能横切面很集中 —— 工具栏在 `src/App.vue`，编辑手感与分页渲染在
`src/components/WordPaper.vue`（约 1300 行），模型/渲染/导出又在 `src/lib/*` 横穿。
任意一组成员功能的文件集都高度重叠，多代理同写同一批文件最后要手工合并，得不偿失。
**默认串行；只有「新增独立纯函数模块」这类真正不重叠的活儿才并行。**

**派发实现代理时，必须写进任务里**：
1. 同步扩 `scripts/` 下对应的验收脚本（这是本仓库的既有规矩：每一条功能都带断言，见 README 的验收表）；
2. 收尾必须跑 `npm run type-check` 与相关 `npm run verify:pX`，**全绿才回报**；
3. **不要 commit**（提交由主 session 在验收通过后做）；
4. **结论单独写进 `.qwen/tmp/wX-impl-result.md`（≤ 40 行）**：改了哪些文件、跑了哪些命令与结果、
   「没做到 / 不确定」的部分（不许美化）；细节与完整日志留在 `wX-impl.log` 里，不要灌进结论文件。

**派发验收代理时，必须写进任务里**（这是踩过坑的，见第 9 节）：
- 仓库内**任何文件都不许改**（发现缺陷只报告不修）；需要临时脚本只许写在 `.qwen/tmp/` 下；
- **长命令输出一律重定向到文件**（`%TEMP%\` 或 `.qwen/tmp/`），只回摘要；
- 结论写进 `.qwen/tmp/wX-review.md`（≤ 40 行：逐条主张给「确认 / 被证伪 / 无法验证」+ 一个总判定），
  证据与命令输出放同一个文件的后面；主 session 只读小窗口，不许把长日志粘进结论；
- 任务是**证伪**实现者的自报，不是复述。
- 不允许自行唤起office word。如果要验证docx文件是否正确，可以交给用户，或者自己解压看xml

**提交风格**：中文 commit message，多条改动用 ①②③ 分段讲清「为什么」，结尾标注对应的 issue。
一次提交装一波。

---

## 1. 写入时状态（2026-09-13）

- **P1 / P2 / P3 已验收**（规格表 → docx、A4 分页预览、编辑层），见 README 的进度表。
- **W1 已完成并提交**（2026-09-13，`4ca96c8`，feat: 编辑器快捷键组、特殊空格与打印）。
  它动了 18 个文件 + 1 个新文件 `src/lib/edit/amount.ts`（下面是提交时的清单）：

  ```
  M README.md                       M src/lib/docx/export.ts     M src/lib/render/css.ts
  M scripts/assert-docx.mjs         M src/lib/edit/dom.ts        M src/lib/render/html.ts
  M scripts/check-docx.ps1          M src/lib/edit/model.ts      M src/lib/spec.ts
  M scripts/test-edit-model.mjs     M src/lib/index.ts           M src/lib/types.ts
  M scripts/verify-docx.mjs         M src/lib/md/parse.ts
  M scripts/verify-editor.mjs       M src/lib/md/serialize.ts
  M src/App.vue                     M src/components/WordPaper.vue
  ?? src/lib/edit/amount.ts
  ```

  W1 内容：`ctrl+U` 下划线（全链路新增 `TextInline.underline`）、`ctrl+shift+E` 修订模式、
  `alt+4` 金额格式化（`#,#00.00`，失败弹提示条）、`ctrl+P` 打印只出 A4 纸、工具栏「插入空格」
  三种特殊空格（U+2003 / U+2002 / U+2005），外加把 `listItem`（列表段落）的首行缩进由 2 字符改成 0。
- 全量 `npm run verify` 在 W1 实现阶段跑过一次并全绿（`%TEMP%\wtp-verify.log` 末尾
  `MARKER_EXIT=0`，五阶段全 PASS）。
- W1 的独立验收结论是**可接受**：七项定向证伪全过 —— 最要紧的一项（修订标记那份由样式表画的
  `text-decoration: underline` 会不会被读回成用户下划线）用浏览器实测确认**无污染**
  （`getComputedStyle` 是 underline 而读回 `underline: false`）；`ctrl+P` 未被拦截；
  打印隐藏清单齐全（含新增的 `.toast`）且 `!important` 的说法成立；`check-docx.ps1` 仍纯 ASCII
  且无 BOM；`formatAmount` 无精度丢失与溢出；`styles.xml` 的 `WT-ListItem` 无脏值；无越界改动。
  两处非阻断的次要项见第 8 节第 4 条。
- **W2 已完成并提交**（2026-09-13，`734dc48`，feat: 文件模板（样式与页边距绑定）与双页并排）。
  它动了 12 个文件、+803/−317：`README.md`、5 个 `scripts/*.mjs`、`src/App.vue`、
  `src/components/WordPaper.vue`、`src/lib/index.ts`、`src/lib/render/css.ts`、`src/lib/spec.ts`、
  `src/lib/render/measure.ts`（只有一段注释）。全量 `npm run verify` 五阶段 PASS
  （`.qwen/tmp/w2-verify.log` 末尾 `VERIFY_EXIT=0`）；独立验收结论**可接受** —— 细节见第 4.5 节，
  次要项见第 8 节第 5 条。

**W2 提交之后工作区是干净的**。如果你打开这个仓库时看到未提交改动，那它属于某个后续波次的在制品：
先 `git status --short` 与 `git diff` 看清是什么，再决定是「补做验收后提交」还是「接着做」；
**不要**直接 `git checkout`／`git stash` 丢掉它。

---

## 2. 不可违背的项目约束（改任何东西前先记住）

1. **正常输入不得触发重排**。渲染只读 `viewDoc` / `pages` 两个浅响应式快照，编辑中的模型不参与渲染；
   每次 `input` 只重量内容变过的块，只有分页结果真的变了才重建 DOM 并把插入符按
   「块 id + 字符偏移」放回去。**新加的任何交互都不要破坏这条**（这是 P3 验收的核心断言）。
2. **规格表是唯一真相源**。`src/lib/spec.ts` 是预览 CSS 与 docx 样式的共同来源；改样式只改那张表，
   预览 CSS（`render/css.ts`）、docx 样式（`docx/export.ts`）、量测段距三处都从它派生。
3. **验收脚本的期望值从 `resolveSpec()` 推导，不硬编码数值**。所以改 spec 之后如果断言没跟着变，
   说明你写错了地方；反过来，样式类改动通常**不需要**改断言。
4. **分页是自研的**（实测行盒 → 装箱 → 在行边界切开），不是 CSS 断页，也不用 paged.js 之类
   （它们靠重建 DOM 分页，会摧毁正在输入的插入符）。表格、并排、节编辑都必须落在这套分页器上。
5. **DOM 是手感的真相，模型是导出的真相**。结构性操作（回车分段、退格合并、改格式、加批注、
   插入表格…）直接改模型再重排渲染；打字、组字、选区、剪贴板留给浏览器。
6. **docx 的内置样式归属按 `w:name` 逐字匹配 Word 本地化名**（含「标题 1」中间那个空格），
   不是按 `w:styleId`。`w:styleId` 一律保留 `WT-` 前缀 —— 撞上 `docx` 库的内置表会让它注入重复定义。
7. `scripts/check-docx.ps1` 必须保持**纯 ASCII**（它对 Word COM 很敏感，非 ASCII 会出问题）。
8. 文档与源码注释一律中文；注释只写「为什么」，不写叙述性内容。

---

## 3. 文件地图（接入点）

> ⚠️ 下面这些行号是 **W1 动手前**测绘的；W1 之后 `WordPaper.vue` / `App.vue` 等都有偏移。
> **用前先 grep 核对**，别照着行号直接改。

| 想改什么 | 落在哪 |
| --- | --- |
| 新增/修改样式（字体字号行距段距缩进对齐编号） | `src/lib/spec.ts` 的 `BlockKind` / `BLOCK_KINDS` / `STYLE_KEYS` / `DEFAULT_SPEC.styles`。`Record<StyleKey, TextStyleSpec>` 会强制补齐，漏了编不过 |
| 页边距 / 页面尺寸 / 模板 | `src/lib/spec.ts` 的 `PageSpec` / `MarginPreset` / `DEFAULT_SPEC.page` / `resolveSpec()` |
| 工具栏按钮（UI） | `src/App.vue` 的 `.toolbar`（约 218-311）。项目里**没有菜单栏/下拉面板**概念，唯一的下拉惯用法是页边距那个 `<select>` |
| 编辑器侧的行为与对外 API | `src/components/WordPaper.vue`：`onKeydown`（约 612-661，**只挂在 `.wtp-content` 上，不是全局监听**）、`defineExpose`（约 999-1030，25 个公开方法） |
| 全局快捷键 | 目前全项目只有一处 `addEventListener`：`document` 的 `selectionchange`。要全局键盘得自己加监听 |
| 选区感知 | `src/lib/edit/dom.ts` 的 `currentRange` / `selectedRanges` / `placeCaret` / `placeRange` / `pointToOffset` / `offsetToPoint`；回显走 `WordPaper.vue` 的 `emitSelection` → 事件 `selection-change` → `App.vue` 的 `selection` |
| 模型操作（纯函数层） | `src/lib/edit/model.ts`：`insertText` / `deleteRange` / `replaceRange` / `splitBlock` / `mergeIntoPrevious` / `setBlockKind` / `applyFormat` / `rangeIsBold` / `rangeColor` / `addComment` / `updateComment` / `removeComment` / `insertBreakAfter` / `removeBreak` / `cloneDoc` … |
| 分页与分节 | `src/lib/render/paginate.ts`：数据结构 63-76、页码推进 `stepNumbering`（约 109）、主循环（约 134）；布局比对 `sameLayout`（WordPaper.vue 内） |
| 量测 | `src/lib/render/measure.ts`：`measureDocument`（带增量缓存）、`MeasureCache`；**换规格表时必须 `clearMeasureCache`** |
| 分页预览的 HTML/CSS | `src/lib/render/html.ts`（行内标记 → HTML，预览/量测/编辑读回共用）、`src/lib/render/css.ts`（规格表 → 预览 CSS） |
| docx 导出 | `src/lib/docx/export.ts`：`groupSections`（分节切组）/ `pageNumberParagraph` / `buildDocument` / `paragraphStyles` / `buildZip`（打包后再补写 `styles.xml` 的 `beforeLines`）/ `toBlob` |
| md 语法 | `src/lib/md/parse.ts`（类 md → 模型）、`src/lib/md/serialize.ts`（模型 → 类 md）。**注意 `verify-docx.mjs` 会断言「模型 → md → 模型」往返一致**，加新语法必须双向可逆 |
| 对外导出 | `src/lib/index.ts` |
| 撤销粒度 | `WordPaper.vue` 的 `Snapshot` / `pushHistory`（栈上限 200） |
| 验收脚本 | `scripts/`，命令映射见 README「验收方式」表 |
| 第三方库能力边界 | `node_modules/docx/dist/index.d.ts`（**只读 d.ts，不要猜 API**） |

---

## 4. W2 —— 样式与页边距绑定（两套模板）+ 双页并排（已完成 2026-09-13）

**需求原文（issue 第 7、4 条）**：

> 7. 当前我们有一套基础样式 + 页边距模板。要求改成：① 样式与页边距绑定。一共有 2 套样式；
>    ② （样式一）当前样式，页边距 25mm，取名为 `管理人文件`；③ （样式二）简易公文格式，页边距使用公文标准。
> 4. 页面宽度足够时，同时展示两页纸。

### 4.1 模板（已与用户定稿，不要再改设计）

- 用户已拍板：**样式二只换页边距，样式一与二的样式值完全相同**。
  - `管理人文件` = 现有 `DEFAULT_SPEC` 全套样式 + 页边距四边 25mm（= 现有 `MARGIN_PRESETS[0]`）
  - `简易公文格式` = **同一套样式** + 公文标准页边距 上 37 / 下 35 / 左 28 / 右 26 mm（= 现有 `MARGIN_PRESETS[1]`）
- 新增 `DOC_TEMPLATES: { key: string; label: string; spec: DeepPartial<Spec> }[]`，模板 = **样式覆盖 + 页边距绑成一体**。
  现在两套的样式值相同，但结构上把位置留出来，将来真的给样式二不同的字体字号时不用改结构。
- `MARGIN_PRESETS` **改为由 `DOC_TEMPLATES` 派生**（不删这个导出，避免破坏已发布的 lib 接口；
  同时保住单一真相源）。
- `src/App.vue` 顶栏那个「页边距预设」`<select>` 换成「文件模板」下拉，默认仍是「管理人文件」
  （保持现有行为，既有断言就不用动）。**样式与页边距绑定之后，再让人单独选页边距是自相矛盾的，所以那个下拉要被取代。**

### 4.2 两个必须处理的坑

1. **切模板必须清量测缓存并整篇重量**。页边距一变，版心宽高就变，`measure.ts` 里所有行数/行高全部失效。
   `measure.ts` 的注释里已写明这条前提（规格表换了一套样式时必须 `clearMeasureCache`），漏了就会
   拿旧量测值算页码。**而且这一条要有断言**：切模板后页数应随之改变（公文标准版心更矮 → 页数通常变多）。
2. **「简易公文格式」也必须过「预览页数 = Word 页数」对账**（`npm run verify:pages`）。
   现在的 `verify:pages` 只验默认模板 —— 公文标准那套版心更矮，正是最容易和 Word 数字对不上的情形。

### 4.3 双页并排

issue 已说清（宽度足够就两页一排），按**自适应**做：容器够宽两页并排，不够自动回到一页一排。
（用现成的 `flex-wrap` 就能自然满足「宽度足够时」这个条件，不要做缩放/适配页宽。）
两个必须收好的地方：

- **页间那枚换页标记要占满整行**（`flex-basis: 100%` 之类），否则并排时会被挤进两页之间错位；
  分页符/分节符的标记是夹在两页之间的 DOM，不在版心内、不参与分页高度计算（这一点别改坏）。
- **打印仍须一页一张**。`@media print` 目前已经是一页一纸（靠 `break-before: page`，第一张回 `auto`），
  并排布局不能把它带偏。注意 SFC 的 scoped 样式带 `data-v` 属性、特异性会压过一层类名，
  必要时用 `!important`（这是 W1 实测撞出来的，见第 9 节）。

### 4.4 涉及文件与验收

`src/lib/spec.ts`、`src/lib/render/css.ts`、`src/App.vue`、`src/components/WordPaper.vue`、`README.md`、`scripts/`。

验收要新增：
- `scripts/verify-browser.mjs`：两套模板各自的样式对账（字号/行高/对齐/缩进/段距）都要跑；
  **并排布局下每页仍不得溢出**；两页并排时页宽/页间距正确；宽度不足时回落到一页一排。
- `scripts/verify-page-count.mjs`：两套模板都要跟 Word 对页数与分节重编号。
- 切换模板后页数变化、插入符与选区不丢（沿用 P3 那套断言）。

### 4.5 结论（2026-09-13）

设计原样落地，**没有偏离 4.1–4.3 节**。实际做法与验收：

- 模板键是 `manager`（管理人文件）/ `govDoc`（简易公文格式）；`DOC_TEMPLATES` 是唯一真相源，
  `MARGIN_PRESETS` 由它派生，公文那组边距数值全仓库只出现一处（`spec.ts`）。
- 4.2 的两条坑都收住了：`props.spec` 的 watcher 现在是「显式 `clearMeasureCache` → 按旧坐标还原插入符/选区
  → `force` 重排」，并且**断言了页数确实改变**（管理人文件 4 页 ↔ 简易公文格式 5 页，两套都与 Word 吻合）；
  `measure.ts` 的缓存注释改成写明「缓存只在同一份规格表下有效」这个不变式（它以前声称「排版宽度全局固定」，
  换模板之后那句话就不成立了）。
- 双页并排：`.wtp-pages` 可换行横排 + 页面 `flex: none`（不缩放纸宽）；换页标记 `flex: 0 0 100%` 占满整行。
  浏览器实测：宽 2200px 两页同排（间距 18px、纸宽未被压缩、页带不横向溢出）、窄 1000px 回落一页一排、
  打印仍一页一张。
- 独立验收结论**可接受**：七项定向证伪全过（含验收方自写 46 项 playwright 断言做 `admin→gov→admin→gov`
  来回切并核对插入符/选区，以及亲自复现 `verify:p2`/`verify:p3`），无功能性缺陷、无越界改动、无放水断言。
  次要项见第 8 节第 5 条。
- 提交：`734dc48`。

---

## 5. W3 —— 查找替换（ctrl+F / ctrl+G）+ 左侧导航窗格（已完成 2026-09-13）

**需求原文（issue 第 3、2 条）**：

> 3. ctrl + F 支持查找功能，ctrl + G 替换功能：① 可选是否使用正则表达式语法（默认不选）；
>    ② 可选范围（当前选中的文本 / 全文）。
> 2. 左侧增加导航窗格。

### 5.1 建议设计（**未经用户过目，属提案**；本波不涉及数据格式，可直接按此实现，有异议再问）

- **匹配算法做纯函数**，放新文件 `src/lib/edit/search.ts`，签名建议
  `findMatches(doc: DocModel, query: string, opts: { regex?: boolean; scope?: { blockId: string; from: number; to: number }[] }): Match[]`，
  `Match = { blockId: string; from: number; to: number }`（模型文字坐标，与 `edit/model.ts` 同一套约定）。
  这样能在 node 里直接单测，不用起浏览器。
- **正则选项**：`new RegExp` 要 `try/catch`，非法模式给提示（复用 W1 已有的提示条机制）；
  默认不勾正则。注意 `g` 标志与零宽匹配会导致死循环，要显式处理（匹配长度 0 时强制前进一位）。
- **范围选项**：「当前选中的文本」= 取当前选区所在的区间（用 `edit/dom.ts` 的 `currentRange`/`selectedRanges`），
  「全文」= 整个 doc。
- **高亮绝不能动 DOM**：用 **CSS Custom Highlight API**（项目里 `keepSelection` 已经用过这一手），
  否则违反第 2 节第 1 条「正常输入不得重排」。
- **替换落到模型**：`replaceRange` → `pushHistory()` → `refreshLayout({ force: true })`；
  全部替换要**从后往前**改，避免前面的替换让后面的坐标失效。
- **导航窗格**：按标题级别（`h1`/`h2`/`h3`）渲染大纲，点击跳到对应段落。
  数据源必须是 `viewDoc` / `viewNumbering`（渲染快照）而**不是**编辑中的 `doc`，
  否则打字会把侧栏带着重渲染。布局上 `.wtp-root` 已经是 flex 横排（批注侧栏就在里面），
  导航窗格可以挂在同一个容器里；注意窄屏/并排（W2）之后的相互作用。

### 5.2 验收

- 单测：新语法/正则/范围内匹配、零宽匹配、从后往前替换的坐标正确性。
- `scripts/verify-editor.mjs`：ctrl+F 打开查找、高亮出现且**不触发重排**、ctrl+G 替换一项/全部、
  非法正则有提示、范围选项生效、导航窗格点击跳转后插入符落在正确位置。

### 5.3 结论（2026-09-13）

设计按 5.1 的提案落地（界面形态由用户另定，见下）；改动 8 个文件 + 2 个新文件
（`src/lib/edit/search.ts`、`src/lib/edit/outline.ts`），共 +1540/−30。

- **界面形态（用户拍板）**：查找替换是**预览区右上角的浮动面板**（拖标题栏可移、`×`／`Esc` 关闭、
  正则默认不勾、范围默认「全文」、Enter 下一个／Shift+Enter 上一个）；导航窗格是**编辑器最左独立一栏**
  （200px、有标题才出现、可折叠、顶栏另有「导航」开关）。
- **高亮**走 CSS Custom Highlight（沿用既有 `keepSelection` 那一手），并且**一个匹配横跨两页时按分页片段
  切成多个子 Range** —— 单个跨页 Range 会把页码与页间换页标记一并圈进去。正文 DOM 一个节点都不改，
  「正常输入不得重排」因此不受影响（有断言：给片段打标记、查找／跳转后标记仍在）。
- **修订模式下的替换**按 Word 语义写「原地标 `w:del` + 其后插入 `w:ins`」两份留痕；且**`del` 标记的文字
  不参与查找与替换** —— 否则「替换一处」会反复命中同一处已删文字、不断叠加（这是收敛的前提，有断言）。
- **打字不带动左栏**：大纲与匹配计数都在 `refreshLayout` 快照更新之后重算，且只在指纹／数字真的变了时
  才 `emit`。
- **验收**：`test-edit-model` 新增第 11–16 节；`verify-editor` 新增 P/P2/P3/Q 四节；全量
  `npm run verify` 五阶段全绿（`type-check` / `verify:p1` / `verify:p2` / `verify:p3` / `verify:pages`），
  另有 `verify:docx` 也过。独立验收代理 14 项定向证伪全过，只有一项判**不可接受**：
  `replaceCurrent` / `replaceAll` 调 `refreshLayout({ force: true })` **没传 anchor**，替换后插入符丢、
  接着敲字会插到段首甚至文档标题最前面 —— 而 `verify-editor` 的 P3 当时只验替换后的模型文字、不验插入符，
  所以绿灯也漏过了。已改为两处都传锚点（`caretAfterReplace`：修订模式下新文字落在被删文字之后，所以从
  `match.to` 起算，否则从 `match.from` 起算），并补三条断言（含「替换后接着敲字落点」的实测）；
  聚焦复核（复核方另写探针专验修订模式那一路）结论**可接受**。
- **本轮教训（已写进第 0 节）**：实现代理那个非交互 session **跑不了任何命令** —— `--approval-mode
  auto-edit` 只自动批准写文件、不批准 shell，所以它交出过「代码写完、命令一条没跑」的成果，类型错误还留在
  里面。凡任务含「跑 npm／起浏览器」一律给 `-y`；主 session 派发后要自己先跑一次 `type-check` 兜底。
- 另外这次顺手查清了一个**既有的分页缺陷**（跨页段落页尾连按 Backspace 会让断点字符重复、越删越多），
  机制与复现探针见第 8 节第 6 条，留待后续波次修。

---

## 6. W4 —— 表格编辑器（拆成 W4a / W4b）

**需求原文（issue 第 5 条）**：

> 5. 增加表格编辑器，表格总宽度默认占满页面（即页边距之内的宽度）；
>    ① 表格的行高需要两个选项：1 是最小两行、最大动态（根据样式`列表段落`的两行）；二是最小一行、最大动态；
>    ② 默认禁止跨页断行；
>    ③ 表格支持增加：(1) 表头行，在 thead 的上方再增加一行，无边框，整行合并，右对齐，用来写`单位：元`这样的文本；
>       (2) 附注行：在 tbody 最下面增加一行，无边框，整行合并，左对齐、顶端对齐，用来写附注。
>       这两个功能都支持使用 radio 来添加/删除；
>    ④ 单元格内部要支持 shift+enter 换行。

**模型设计已于 2026-09-13 给用户过目并拍板。完整记录在项目记忆 `project/table-model-design.md`，
下面是要点复述；动手前请以记忆文件为准。** 这是本计划里唯一一个已经过闸门的波次。

### 6.1 已拍板的设计（勿再改回）

- **单元格 = 可寻址的「伪块」**：格子 id 用 `cellId(tableId, r, c)` → `` `${tableId}.r${r}c${c}` ``，
  `<td>` 内那层 div 挂 `data-block-id`。这样 `edit/dom.ts` 的
  `fragmentOf` / `pointToOffset` / `offsetToPoint` / `placeCaret` / `placeRange` / `currentRange`
  **一行都不用改**（`fragmentOf` 向上取最近的 `data-block-id`，格内光标自然落到格子上）。
  代价是 `edit/model.ts` 的 `findBlock(doc, id)` 要泛化成 `findContainer(doc, id): { inlines: Inline[] } | undefined`，
  `insertText` / `replaceRange` / `deleteRange` / `applyFormat` / `rangeIsBold` / `rangeColor` 改吃容器。
  **已否决**给坐标系加第二层（`blockId + row + col + offset`）的方案（那要改遍 `dom.ts`）。
- **断行按行级**：单行不拆，表格可在行与行之间断开续到下一页（Word 的 `w:cantSplit` 语义）。**不是整表不拆**。
- **Shift+Enter 走真软换行 `<w:br/>`**（Word 语义），单元格保持单段落。需要给 `Inline` 加一个软换行标记，
  量测 / 坐标映射 / md 序列化都要跟着认它。
- **表内文字直接复用「列表段落」样式**，不新增「表格文字」样式。
  （W1 已把 `listItem` 的首行缩进改成 0，正是为了这件事 —— 所以**不要**再在任何地方加单元格专用的缩进豁免。）
- **列标题行不独立表达**：表格里所有行结构上都是普通行，`role` 只剩 `unit` / `body` / `note`
  （用户要加粗居中的列标题就自己套格式）。因此 Word 那种「跨页自动重复标题行」不做。

### 6.2 模型形状

```ts
export type TableRowRole = 'unit' | 'body' | 'note'

export interface TableCellModel { inlines: Inline[] }            // 单段落
export interface TableRowModel  { role: TableRowRole; cells: TableCellModel[] }

export interface TableBlock {
  t: 'table'
  id: string
  rows: TableRowModel[]     // 顺序即显示顺序：unit → body… → note
  columns: number           // 取 body 行的最大格数
  minLines: 1 | 2           // 「最小一行 / 最小两行」，即 HeightRule.ATLEAST 的倍数
  cantSplit: boolean        // 默认 true
}

export type Block = TextBlock | SectionBreakBlock | PageBreakBlock | TableBlock
export function cellId(tableId: string, r: number, c: number): string
export function parseCellId(id: string): { tableId: string; row: number; col: number } | null
```

- 「整行合并」不存字段：`unit` / `note` 两个 role **天然就是整行一格**，渲染与导出时展开成
  `columnSpan = columns`。这样合并态不可能和实际格数不一致。
- 行高值 = `minLines × 「列表段落」的行高`（该样式是 `lineRule: 'atLeast'`、`linePt: 12`）。

### 6.3 四个接缝怎么改

| 接缝 | 改法 |
| --- | --- |
| 量测 `render/measure.ts` | 表格走独立分支：**只量每行的高度**（`<tr>` 实测高），不量 `rowStarts`（行是原子的，不需要行内切分） |
| 分页 `render/paginate.ts` | 表格按**行**装箱：`MeasuredItem` 加一种 `tableRow { blockId, row, height }`，每项 `rows=1`、`lineHeight=height`；**必须加 `atomic` 标志绕开孤行控制**（否则单行块会被 widow/orphan 判成「不足 2 行」整块挪走）；同页相邻同表行合并成一个片段，`PageFragment` 加 `rowFrom`/`rowTo`（`from`/`to` 那套字符偏移对表格无意义）；一页一个 `<table>` |
| 渲染 `render/html.ts` + `render/css.ts` | 新增 `renderTableFragment(block, rowFrom, rowTo)`，**量测与预览必须共用同一个函数**（这是现有量测可信的前提）；表格宽度取版心宽；行高最小值用 CSS `min-height` |
| 导出 `docx/export.ts` | 见下 |

### 6.4 docx 导出映射（`docx` v9.7.1 实测，坑都标出来了）

- 表格：`new Table({ rows, columnWidths, layout: TableLayoutType.FIXED, borders: {...} })`。
- **表格总宽度**：`width: { size: <版心宽缇值>, type: WidthType.DXA }` —— **`type` 必须显式写**，
  库的默认是 `AUTO` 而不是 `DXA`（`node_modules/docx/dist/index.mjs:20432`：
  `var createTableWidthElement = (name, { type = WidthType.AUTO, size }) => {`，此条已实测核对），
  不写就落到 `auto`；另外 `w:tblW` 只在 `options.width` 存在时才写出（同文件 21294 行）。
  版心宽 = 页面宽 − 左右页边距，需要从 `spec.page` 换算成缇。
- 行：`cantSplit: true`（**只在 `TableRow` 上有这个字段**）；
  `height: { value: ptToTwips(minLines × 列表段落行高), rule: HeightRule.ATLEAST }`（ATLEAST 才是「最小 N 行、最大动态」）。
- `unit` / `note` 行：一格 + `columnSpan = columns` + 单元格 `borders` 全 `NONE`；
  **单元格没有水平对齐字段**，右对齐/左对齐只能落在格内 `Paragraph.alignment`；附注行的顶端对齐用
  `TableCell.verticalAlign = TableVerticalAlign.TOP`。
- 软换行：`TextRun.break?: number` 或 `CarriageReturn` 类（本项目此前只用过 `PageBreak`）。
- `Table` 与 `Paragraph` 可以混排进 `sections[].children`。
- 已知未做：跨页重复标题行（模型里没有表头概念）。

### 6.5 md 语法（提案）

新增围栏块，避免和现有「一行一段」的规则打架：

```
:::table minLines=2
> 单位：元
| 项目 | 金额 |
| 甲 | 1,234.00 |
< 注：以上金额不含税
:::
```

`>` 开头是 unit 行、`<` 开头是 note 行、`|` 开头是数据行。
**必须双向可逆**（`verify-docx.mjs` 会断言「模型 → md → 模型」往返一致）。

### 6.6 W4a / W4b 拆分

> **2026-09-13 更新：W4a 又拆成 W4a-1（已完成并提交）/ W4a-2（待做）—— 结论与新的验收形态见 6.7；下面这段保留原始拆分理由。**

- **W4a（先做）**：模型 + 渲染 + 量测 + 分页 + docx 导出 + md 语法。
  目标：**表格排得出来、导得出去、Word 对账通过**。此阶段可编辑性最小化。
  验收：`scripts/verify-docx.mjs` 加含表格的样本；`check-docx.ps1` 用 Word COM 读回表格结构
  （行数/列数/合并/行高规则/边框/对齐）与规格表对账；`test-paginate.mjs` 加「表格行按行装箱 +
  atomic 不吃孤行控制」的单测；`verify:pages` 用含表格的样本比对页数。
- **W4b（后做）**：编辑交互 —— Tab / 方向键跨格、加行/加列/删行/删列、radio 加删 unit/note 行、
  行高两档切换、格内 Shift+Enter。
  验收：`verify-editor.mjs` 加浏览器实测（格子内输入不重排、跨格移动、增删行列后模型与 DOM 都正确）。

**理由**：表格最大的风险在「导出的 docx 与预览版式对不上」，这部分能独立验收；交互是纯前端增量，
放第二步不会拖累第一步的验收。

---

### 6.7 W4a-1 结论（2026-09-13，已提交）

**W4a 又拆成两小步**：**W4a-1 = 表格的模型 + md 围栏语法 + docx 导出 + Word 读回对账**（不碰预览/编辑/demo）；
**W4a-2 = 预览渲染 + 量测 + 分页 + 编辑读回收口 + demo 样本放表 + 工具栏「插入表格」按钮 + 软换行**
（`Inline` 的 `br` 标记、md 的 `\n` 写法、`<w:br/>` 归 W4a-2 —— 它的编译波及面全在编辑侧）。
用户同日拍板：W4a 就带「插入表格」按钮，且**格内可以直接打字**（靠 `edit/model.ts` 的 `findBlock` → `findContainer`）。

**实际改动 11 个文件**：`types.ts`（`TableBlock`/`TableRowRole`/`cellId`/`parseCellId`）、
`edit/model.ts`（**只有** `cloneDoc` 一处深拷贝）、`md/parse.ts` + `md/serialize.ts`（`:::table` 围栏与往返）、
`docx/export.ts`（表格映射）、`index.ts`（导出）、`scripts/{verify-docx,assert-docx,check-docx.ps1,test-edit-model}.mjs`、`README.md`。

**验收走了两轮**：第一轮独立验收判**不可接受** —— `splitTableCells` 只看反斜杠、不跳 `{}` / `[[]]`，
格内 `{红|甲}`、`[[甲|核对原件]]` 的竖线被当成列分隔符（一格拆多格、颜色与批注静默丢失）；另有未闭合围栏吞掉后文、
unit/note 行尾空格丢失。修完后聚焦复核判**可接受**，且复核方用**变异测试**证明新用例真能抓到这个缺陷
（把切格算法回退成旧版 → 9 项变红）。顺带修了一处**与表格无关的既有缺陷**：`[[锚|内容]]` 的**批注内容**里含 `|` 或 `\` 时，
序列化会逃逸、解析却不还原，每往返一次就多一层反斜杠（不收敛）。
最终：`type-check` 0 错、`test-edit-model` 239/239、`test-paginate` 67/67、`verify:docx` 全绿、
`verify:p1` 的 Word 逐项对账 `[PASS]`（含新增的表格节：行数 / 格数 / 整行合并 / 行高规则 / 边框 / 对齐 / 总宽）。

**W4a-2 的验收形态必须改（2026-09-13 决定）**：本机 Edge 的程序化启动坏了（见 9.7），用户要求
**不跑 playwright、只做代码级验证** —— 所以「预览页数 = Word 页数」这类机器对账做不了。
机器只验能得到确定答案的部分（分页器单测、md 往返、docx 字节层），版式一致性改成**人工核对**
（开 dev server 看页面 + Word 打开导出的 docx 对比页数与表格外观）。

**列标题行居中仍做不到**（见 8.2），本轮样本的列标题行只套了 `**加粗**`。

### 6.8 W4a-2 开工要点（已完成 2026-09-13 —— 结论见 6.9）

**范围**（用户 2026-09-13 拍板：带 UI 入口、格内可打字）：预览渲染 + 量测 + 分页 + 编辑读回收口 +
demo 样本放一张表 + 工具栏「插入表格」按钮 + 软换行。**W4b 只剩** Tab / 方向键跨格、增删行列、unit/note 行开关、
行高两档、格内 Shift+Enter 键入。
**模型与 md / docx 侧已就绪**（W4a-1，提交 `d2911f7`）：`TableBlock` / `cellId()` / `parseCellId()`、
md `:::table` 围栏、docx 映射全在 —— **不要重新设计**（设计源头见 6.1–6.5 与项目记忆 `table-model-design.md`）。

**已知接缝与坑（动手前先看）**：

- `render/html.ts`：新增 `renderTableFragment(block, rowFrom, rowTo)`，**量测与预览必须共用这一个函数**；
  单元格内层 div 挂 `class="wtp-listItem"` + `data-block-id="cellId(...)"` + `data-from/data-to`
  （这样 `edit/dom.ts` 的坐标换算一行都不用改）。
- `render/css.ts`：单元格左右内边距必须与 docx 侧一致 —— **108 缇 = 5.4pt**（导出侧 `CELL_MARGIN_TWIPS` 已定死）。
  行高最小值**不能靠 `<tr>` 的 `min-height`**（表格行不吃这个属性）：要么给 `<tr>` 写 `height`（表格布局里按最小值处理），
  要么给格内 div 设 `min-height` —— 必须浏览器实测确认。已知偏差：列表段落样式的 CSS `line-height: 12pt` 是固定值，
  而 Word 的 `atLeast 12pt` 是「不小于」；单行格两者都是 24pt（`minLines=2`）没问题，**多行格**（内容撑高）
  可能差零点几磅 —— 页数对不上时先量这里。
- `render/measure.ts`：`MeasuredItem` 加「表格行」一项（每行 `rows=1`、`lineHeight=行高`），量测缓存条目类型放宽成联合，
  签名要含整张表的 HTML（任一格改动即失效）；表格**按行**量测（`<tr>` 实测高）。
- `render/paginate.ts`：表格行带 `atomic` **绕开孤行控制**（否则单行块会被 widow/orphan 判成「不足 2 行」整块挪走）；
  `PageFragment` 加 `rowFrom/rowTo`；同页相邻同表行合并成一个片段（一页一个 `<table>`）。
- `components/WordPaper.vue`：片段渲染加表格分支（`blocksById` 之外要有表格查找表）；`fragmentStyle` / `sameLayout`
  要比 `rowFrom/rowTo`；**v-for 的 key 必须带 `rowFrom/rowTo`**（同一张表各页片段的 `from/to` 都是 0，只按
  `(blockId, from)` 编键会撞）；编辑读回收口把 `findBlock` 泛化成 `findContainer`
  （`insertText`/`replaceRange`/`deleteRange`/`applyFormat`/`rangeIsBold`/`rangeColor` 改吃容器），
  `edit/search.ts` 同步改吃容器，好让查找/替换覆盖格内文字。
- **软换行**（本波唯一的新数据格式）：给 `Inline` 加零宽标记（Word 的 `<w:br/>`），md 里写成 `\n`（反斜杠 + n，
  与 `escapeText` 的 `\\` → `\` 可逆），docx 用 `TextRun.break` / `CarriageReturn`；
  **量测、坐标映射、`replaceRange`、`sliceStrict`、`readInlines`、`sliceInlines` 都要认它**（零宽、不占字符位）；
  `readInlines` 要能区分「空段落占位用的 `<br>`」与真软换行（只在根节点仅有一个 `<br>` 时算占位）。
- **demo 样本放表**会动 `verify-browser` / `verify-editor` 里依赖 demo 内容的断言（搜索命中数、打印页数、切模板页数…），
  要一起改；表格文字别用含下划线 / 修订 / 批注的写法，免得撞上既有计数断言。

**验收形态（用户 2026-09-13 拍板：不跑 playwright，版式改人工核对）**：

- 机器验：`type-check`；`test-paginate.mjs`（表格行按行装箱 + `atomic` 不吃孤行控制）；`test-edit-model.mjs`
  （软换行 / 容器泛化 / 跨格查找）；`verify:docx`（`<w:br/>` 字节层 + 往返）；`verify:p1`（Word 对账，
  样本里加一个软换行格）。
- **人工核对**（用户做，届时给一份清单）：开 dev server 看表格版式（边框、对齐、行高、合并行、列宽、分页处行不拆开），
  再用 Word 打开导出的 docx 对页数与表格外观。

**派发方式（用户 2026-09-13 拍板：完成通知不要再进对话）**：子进程**完全脱离**，不注册成 harness 的后台任务 ——
把启动命令写成 `.qwen/tmp/wXn-run.cmd`（`call qwen -p "Read .qwen/tmp/wXn-task.md as UTF-8 and execute it exactly as written." -y > .qwen/tmp/wXn-impl.log 2>&1`；
任务书本身写成 UTF-8 文件，`-p` 只给一句纯 ASCII 指针），再用
`Start-Process -FilePath cmd.exe -ArgumentList '/c','.qwen\tmp\wXn-run.cmd' -WindowStyle Hidden` 起。
主 session 用**前台等待循环**（每次 ≤9 分钟，轮询结论文件是否出现）取结果，读文件一律有界。
代价：失败 / 挂住没有信号，只能靠等待超时 + 读日志尾部发现 —— 所以轮询循环里要同时判
「结论文件是否出现」与「进程是否还在」，两者都不满足才退出。

⚠️ **W4a-2 实测订正两条**：
① **`start /b cmd /c xxx.cmd` 不行**。`start /b` 共享控制台，shell 工具会把整条命令等满超时（默认 120 s），
而超时又按进程组把子进程一起 `^C` 掉（日志里只剩一句 `^C`）。必须用 `Start-Process … -WindowStyle Hidden`，
它的子进程不在工具的进程树里，工具立刻返回、也不弹窗。
② 这条命令会被本 session 的 `auto` 批准模式**拦一次**（判为「隐藏窗口 + 脚本间接调用 `qwen -y` = 绕过权限判定」）。
按规则**不换路径绕、原样重发同一条命令**即走人工批准，放行后正常脱离运行。
另外，子进程在 `-p` 模式下的 stdout 要等最终产出才落盘（W4a-1 一整轮日志也只有 3.4 KB），
**「日志没长」不等于进程死了** —— 判活用 CPU 时间或仓库内文件的 mtime 更靠谱。

---

### 6.9 W4a-2 结论（2026-09-13，已提交）

**设计原样落地，但有一处 PLAN 自己的提议被证伪**：软换行的 md 写法改用**方案 A（`{br}` 花括号指令）**。
6.8 原提的 `\n`（反斜杠 + n）**不可行** —— `md/parse.ts` 的 `parseInline` 第一条分支就是「反斜杠吃掉下一个字符」，
`\n` 会被吃成字面量 `n`（要救就得把判断插在那条通用分支**之前**）。实现者改用 `{br}`，加在 `parseDirective`
的「没有竖线就当字面量」之前；正文里真写「{br}」会被 `escapeText` 写成 `\{br\}`，两侧对称、不与既有指令撞车。

改动 **22 个文件**（含本文件 `PLAN.md` 的回填；准确行数见提交）：`src/lib/types.ts`（零宽 `BreakInline`）、
`md/{parse,serialize}.ts`、`docx/export.ts`（`<w:br/>`）、`render/{html,css,measure,paginate}.ts`、
`edit/{model,search,dom}.ts`、`src/lib/index.ts`、`components/WordPaper.vue`、`App.vue`、`README.md`、
6 个 `scripts/*.mjs`。

**本波定下的接口事实（后续波次务必照做）**：

- **表格片段的外层不挂 `data-block-id`**，只挂 `data-table-id` / `data-row-from` / `data-row-to`；
  格内那层 div 才挂 `data-block-id = cellId(...)` + `data-from` / `data-to`。原因是 `edit/dom.ts` 的
  `fragmentOf()` 向上取「最近的 `data-block-id`」——外层若也挂，`retagFragments` / `syncPlain` 的
  `querySelectorAll('[data-block-id]')` 就会把外层元素当成片段。
- 片段渲染的 **v-for key 必须带 `rowFrom` / `rowTo`**（同一张表各页片段的 `from` / `to` 都是 0，会撞键）。
- **行高最小值落在格内 div 的 `min-height`**（`<tr>` 不吃 `min-height`）；`minLines=2` 走 `.wtp-table-min2`。
- 单元格左右内边距 **5.4pt = 108 缇**，与 `docx/export.ts` 的 `CELL_MARGIN_TWIPS` 同源。
- `findBlock` 的对外导出**保留**（只找 textBlock），新增 `findContainer` 吃容器（段落或格子）；
  查找 / 替换 / 读回 / 改格式全走容器，所以格内文字也在查找替换范围内。
- **「插入表格」是浮动规格面板**（与查找替换面板同一套：可拖、`×`／`Esc` 关；默认贴预览区左上，
  查找面板贴右上，两个浮层同时开也不叠），用户在面板里选**行数与列数**，默认 2 行 × 3 列；
  数值不设信任 —— 组件侧夹回 **1–30 行 / 1–12 列**（面板的 `min`/`max` 只是提示，数字是能直接敲的）。
  `minLines=2`、`cantSplit=true`；插到当前块之后（落点在格内就插到那张表之后），插入符放第一个格子开头。
  **增删行列留 W4b。**
- 浮动面板的拖动逻辑收敛成一套（`beginPanelDrag` + `onPanelDrag` + `endPanelDrag`），
  两个面板只各自提供一个 6 行的起始函数；面板打开时把焦点交给第一个输入框
  （数字立刻能敲、`Esc` 才有接收者），焦点离开正文**不会丢插入点** ——
  组件的 `lastCaret` 记着上一次落点（`onSelectionChange` 在选区不在片段里时直接返回、不清它）。
- 格内 **Enter / Shift+Enter 本波只 `preventDefault`**，按键绑定留给 W4b。

**验收**（机器验全绿）：`type-check` 0 错；`test-paginate` **96/96**（原 67）；`test-edit-model` **288/288**（原 239）；
`verify:docx` 全 ok（含软换行字节层）；`verify:p1` Word COM `[PASS]`（含「格内文字里读到 `chr(11)`」与表格对账）。
**版式一致性没做机器验证**（本机 Edge 坏了、按用户 2026-09-13 的决定不跑 playwright），改**人工核对**；
`verify:p2` / `p3` / `pages` 一律没跑。

独立验收结论**可接受（附一处提交前必修缺陷）**：16 项定向证伪里 14 项确认、1 项「基本确认但不是结构性避开」、
1 项**被证伪** —— `scripts/verify-browser.mjs` 新增的表格量测对账用 `rendered.get(it.blockId)` 建键，
而表格片段外层只有 `data-table-id`，于是 `rendered.get(tableId)` 恒为 0 高、断言恒假（方向是 fail-closed，
但一旦有人能跑浏览器就会红）。**已在提交前修掉**：`analyzePage` 的 item 增加 `tableId`，建键改用 `tableId || blockId`。
验收方另自写 **87 项独立探针**全过，并复核 `verify-editor.mjs` 那条「页数 +1 → 页数不减少」属
**论证成立的修法而非放水**（`页数 + 1` 不是恒真：页尾余量被吸收时总数可以不变），但它是推理驱动、**没跑过**。

**验收之后的两处改动（2026-09-13，仍在同一波内）**：① 修掉验收方指出的 `verify-browser.mjs` 建键缺陷
（`analyzePage` 的 item 增加 `tableId`，建键改用 `tableId || blockId`）；② 用户拍板把「插入表格」由
固定 3 列 × 2 行改成**选规格的浮动面板**（见上），并同步补了 `verify-editor` 的 U 节、把 `.table-panel`
加进打印隐藏清单与打印隐藏断言。这两处**都发生在验收之后**，因此又单独派了一轮**增量复核**：
结论**可接受** —— 复核方用 blob 哈希比对确认验收之后只动过 6 个文件（其余 16 个与验收时逐字节相同），
逐条确认了建键修复、拖动收敛无回归、焦点兜底推理成立（`caretPoint() = selectionRange()?.start ?? lastCaret`，
且 `onSelectionChange` 在选区不在片段里时直接 return、**不清 `lastCaret`**）、`clampCount` 的
NaN/空串/0/负数/小数/超大值全部落在界内、U 节断言不放水也不误报、打印隐藏的两组断言自洽、
两面板默认位置（左上 / 右上）在 482px 以上不重叠、无越界改动。
**仍然没做机器验证的**：表格量测对账的**具体数字**、面板交互手感、打印时浮层真的不出现 ——
这三样都要浏览器，`verify:p2`/`p3` 依旧是「写了没跑」。

---

### 6.10 W4b 设计（2026-09-13 拍板，勿再改回）

**用户拍板四条**：① 控件形态 = **上下文工具条**（不是浮动面板）；② Tab 走到**最后一个格子时按了不作任何响应**
（不换行、不自动加行）；③ 本波**拆两轮**：W4b-1（模型操作 + 工具条 + radio + 格内 Shift+Enter + 边界护栏）
→ W4b-2（Tab / 方向键跨格）；④ 派发子进程**不钉 `--model`**（用 CLI 默认，主 session 自己跑 type-check 兜底）。

**模型不动**：`TableBlock` 已能表达全部状态（`rows` 顺序即显示顺序、`role`、`columns`、`minLines`、`cantSplit`）。
增删行列 = 数组增删 + `normalizeTable()` 重算 `columns` + 重排渲染；量测缓存签名含整张表 HTML，改一格即失效。

**W4b-1 内容**

- 新文件 `src/lib/edit/table.ts`（纯函数、node 可单测）：`findTable(doc, id)`（id 可以是表格 id 或格子的
  cellId）、`insertBodyRow(table, at)`、`removeBodyRow(table, at)`（body 行只剩 1 行时拒绝）、
  `setRoleRow(table, 'unit' | 'note', on)`（各至多一行）、`insertColumn(table, at)` / `removeColumn(table, at)`
  （只作用于 body 行，`columns` 会掉到 0 时拒绝）、`normalizeTable(table)`（`columns` = body 行最大格数且 ≥1；
  unit/note 行只留第 0 格）、`setMinLines()`。
- `EditorSelection`（`edit/model.ts` + `index.ts` 导出）增加可选 `table` 上下文：tableId / row / col / role /
  rows / columns / bodyRows / minLines / hasUnit / hasNote，由 `WordPaper.vue` 的 `emitSelection()` 从模型现算
  —— App 侧拿不到响应式的模型，上下文工具条只能吃这一次 emit。
- `App.vue`：主工具栏下方多加一条 `v-if="mode === 'edit' && selection?.table"` 的上下文工具条：行（上方插入 /
  下方插入 / 删除行）、列（左侧插入 / 右侧插入 / 删除列）、行高两档 radio、表头行(unit) 有/无 radio、
  附注行(note) 有/无 radio。按钮一律 `@mousedown.prevent`（焦点不离开正文，落点与 native 选区才保得住）。
  禁用规则：删除行（光标在 unit/note 行、或 body 行只剩 1 行）、删除列（`columns <= 1`）。
  **列的下标参照**：unit/note 行天然只有一格（恒为 `c0`），在那两行做列操作就按第 0 列算。
  记得加进 `@media print` 的隐藏清单并同步打印隐藏断言（`.table-panel` 已有先例）。
- `WordPaper.vue` 新增并 `defineExpose`：`insertTableRow(where)`、`removeTableRow()`、
  `insertTableColumn(where)`、`removeTableColumn()`、`setTableMinLines(n)`、`setTableRoleRow(role, on)`。
  每个都是「取落点 → `pushHistory(caret)` → 改模型 → 用**新下标**重算锚点 → `refreshLayout({anchor, force:true})`」。
- **`cellId` 里嵌的是行/列下标**（`tb1.r2c1`），所以任何增删都会让其后的格子 id 整体位移 —— 新锚点必须按
  新下标重算。这是 W4a 定的坐标方案的必然结果，不是缺陷。
- 格内 **Shift+Enter** 插 `{ t: 'break' }`。**落点是个坑**：软换行零宽，`offsetToPoint` 会把同一偏移还原到
  「换行之前」（上一行末尾），照原样 `placeCaret` 会让回车后敲的字打回上一行。做法：`edit/dom.ts` 新增
  `placeCaretAfterBreak(root, point)`（定位到该字符偏移处的 `.wtp-br`，用 `range.setStartAfter(br)`），
  `refreshLayout` 的选项加 `afterBreak?: boolean`，只在 Shift+Enter 这一条路径上传 `true`。
  **不要**去改 `offsetToPoint` 的通用约定（那会连带影响普通段落里已有的软换行）。
- 格内**边界护栏**（现在缺，必须补）：格首 Backspace、格尾 Delete 都要 `preventDefault` —— 否则浏览器原生
  会把相邻 `<td>` 的 DOM 并掉，直接破坏表格结构。格内 Enter 仍不接管（单元格是单段落，见 6.1）。
- 验收脚本：`test-edit-model.mjs` 补 `edit/table.ts` 的结构操作单测（含两处保底：最后一行 body 不可删、
  最后一列不可删；以及 `normalizeTable` 的 columns 归一）；`verify-docx.mjs` 扩样本（带 unit/note 行的多行表）
  与对账；`verify-editor.mjs` **写**上下文工具条与 Shift+Enter 的浏览器断言（本机跑不了，见 9.7）。

**W4b-2 内容**（等 W4b-1 验收完再派）：Tab / Shift+Tab 行优先跨格（unit/note 行整行算一格）；Tab 在最后一格
= 无响应；→/← 在格首/格尾跨格；↑/↓ 在格内首/末视觉行跨格（用 `getClientRects()` 比较插入符与格内首字符 /
末字符的 top 来判定）。纯手感，只能人工核对。

### 6.11 W4b-1 结论（2026-09-13，已提交）

**范围**：新增 `src/lib/edit/table.ts`；`EditorSelection` 加表格上下文；`WordPaper.vue` 增加 6 个表格操作
并 `defineExpose`；格内 Shift+Enter；格首 Backspace / 格尾 Delete 护栏；`App.vue` 的上下文工具条。
**Tab / 方向键跨格按用户拍板留给 W4b-2**（且「Tab 停在最后一个格子时按了不作任何响应」）。

**本波定下的接口事实（W4b-2 与 W5 照做，别再重新设计）**：

- `edit/table.ts` 导出 `findTable` / `bodyRowIndexes` / `bodyInsertIndex` / `insertBodyRow` / `removeBodyRow` /
  `setRoleRow` / `insertColumn` / `removeColumn` / `normalizeTable` / `setMinLines`；**组件层不要再自己实现一份**
  （W4a-2 时组件里那份本地 `bodyIndexes` 已删掉）。
- **插入行的落点必须先过 `bodyInsertIndex(table, at)`**：它把「想插在第几行」夹进 body 区间。不夹的话，
  光标停在 unit 行时点「上方插入行」会把 body 行插到表头之前（note 行对称），破坏 unit → body… → note 的显示顺序。
  锚点行 = `at <= 光标行 ? 光标行 + 1 : 光标行`。
- `EditorSelection.table`（`TableSelectionContext`）由 `emitSelection()` 从模型现算 —— App 侧没有响应式模型，
  上下文工具条只能吃这一次 emit。
- **每个结构操作在 `refreshLayout` 之后都要补一次 `void nextTick(emitSelection)`**：点按钮 / radio 会把焦点
  从正文拿走，`selectionchange` 未必再触发；不补这一下，工具条的 radio 选中态与行列提示会停在旧值
  （用户看着就是「点了没反应」）。
- 格内 Shift+Enter 的落点走 `edit/dom.ts` 新增的 `placeCaretAfterBreak()` + `RefreshOptions.afterBreak`。
  原因：软换行零宽，同一字符偏移既能是「换行前」也能是「换行后」，通用 `offsetToPoint` 一律还原到换行前，
  于是回车后敲的字会打回上一行。**`afterBreak: true` 全仓库只允许 Shift+Enter 这一条路径上传**。
- 格内边界键一律 `preventDefault`：格首 Backspace / 格尾 Delete 若放给浏览器，原生会把相邻 `<td>` 的 DOM 并掉，
  直接破坏表格结构。格内 **Enter 仍不接管**（单元格是单段落）。

**机器验收（全绿）**：`type-check` 0 错；`test-edit-model` **341 → 365**；`test-paginate` 96/96；
`verify:docx` 全 ok；`verify:p1` Word COM `[PASS]`（含「unit + body×5 + note」共 7 行的逐项对账）。
`verify-editor.mjs` 新增 V 节（工具条出现/消失、插入行、删列禁用、两组 radio、Shift+Enter 的落点在换行之后、
格首 Backspace 后模型不变）**写了但本机没跑**（Edge 坏，见 9.7）；`verify:p2` / `p3` / `pages` 一律没跑。

**独立验收与增量复核**：第一轮判**「可接受」**，两条非阻断项 —— ① 在 unit/note 行上插入行会破坏行序；
② 单测没覆盖「**列下标 0 + 含 unit/note 行**」这个「unit/note 豁免」唯一会暴露的角落（第一轮变异 M3/M4
当时红不了）。两条都在提交前修掉：新增 `bodyInsertIndex` 夹取 + 补两组单测；复核方用变异实测确认现在会变红。
提交前主 session 还误跑过一次 `prettier --write`（见 9.4），已逐字节还原，并就此单独派了一轮增量复核，
结论同为**「可接受」**。

---

## 7. W5 —— 节编辑框架

**需求原文（issue 第 1 条）**：

> 1. 增加一个`节编辑`的框架：① 根据指针位置定义所在节；② 在所在节中，可以选择是否增加页码、
>    页码是否关联前节（默认是）、是否从 1 开始编码（默认否）；③ 可以选择当前节的页面方向。

### 7.1 前置：必须先扩模型（**而且必须先出设计给用户过目**）

现状（测绘结论）：

- `Spec.page` 是**全局单份**（`resolveSpec()` 一次浅合并），`export.ts` 里每一节都读同一个 `spec.page`；
  **没有任何按节的覆盖机制**。
- `SectionBreakBlock`（`src/lib/types.ts`）只带一个字段 `restartNumbering: boolean`；
  `paginate.ts` 里分节 = `sectionIndex += 1` + 可选的页码重排，节总是换页。
- 所以「逐节改页面方向 / 页码关联 / 页码起始」在当前模型下**表达不出来**。
  需要把 page 覆盖挂到分节符上（例如给 `SectionBreakBlock` 加一个 `page?: DeepPartial<PageSpec>`），
  并让 `paginate` 的分节步进、`export.ts` 的每节 `properties.page`、预览的每页版心都按节取。

### 7.2 docx 侧的映射（`docx` v9.7.1 实测）

- 页面方向：`properties.page.size.orientation = PageOrientation.LANDSCAPE`（有 landscape）。
- 页码从 1 重编：`properties.page.pageNumbers = { start: 1 }`（`w:pgNumType`）；
  现有代码已经这么做了（只在 `restartNumbering` 时传）。
- **「链接到前一节」在 docx 里没有独立字段**，是靠「不写对应引用」表达的：不写 `pageNumbers`
  就延续前一节的页码序列，不写 `footers` 就沿用前一节的页脚。
  **注意现状是个坑**：`export.ts` 目前对**每一节都无条件写 `footers.default`**，
  等于每节都自带页脚引用、永远不继承。要做「页码关联前节」就必须改成**按节决定是否写 `footers`**。
- 节类型：`properties.type = SectionType.NEXT_PAGE`（默认行为就是新起一页）。

### 7.3 验收

- 单测（`test-paginate.mjs`）：分节 + 页码重排 / 连续编号 / 每节不同版心（横竖混排）的页数计算。
- `verify:p1`：Word COM 读回每节的页面尺寸/方向、`w:pgNumType`、页脚是否存在，与模型对账。
- `verify:pages`：横竖混排 + 页码重排的样本，预览页数必须等于 Word 页数。

---

## 8. 遗留未决问题（不属任何一波，但别忘）

1. **段间距「相加还是取大」尚未钉死**。预览与分页现在把相邻两段的间距算作「段后 + 段前」，
   但 2026-09-13 在真实 docx 上用 Word COM 量首行纵向位置差，5 组结果都指向 Word 其实**取两者较大者**
   （段前段后同为 1 行时两段间距仍是 1 行；段后 3 行 + 段前 1 行时只跟随 3 行那个）。
   若成立，预览的段间距会比 Word 大一倍。**证据链不闭环**（想用直接格式做对照时 Word 没吃我写进
   `<w:pPr>` 的 `w:spacing`；从零建文档的对照也不成立），要动这一块之前先把它钉死。
   详见项目记忆 `project/word-paragraph-spacing-open-question.md`。
2. **表格列标题行居中做不到**：本项目**段落对齐是样式级**的（`TextStyleSpec.align`），
   没有逐段对齐覆盖。用户若想在 W4b 之后要「加粗居中的列标题行」，需要二选一：
   给段落/行加一个 `align` 覆盖字段，或新增一个居中样式。**到 W4b 时问用户。**
3. **编辑层既有取舍**（README 已记）：跨段落选中删除不按 Word 语义合并段落；
   软换行此前不支持（W4 只为单元格引入，之后是否推广到普通段落待定）；
   关掉修订模式后在刚被标为 `w:ins` 的文字后继续输入，新字会并进那个 `w:ins` 区间
   （要修得做逐字 diff，有回归风险，W1 已记入 README 的「已知取舍」）。
4. **W1 验收发现的两处次要项**（非阻断，记在这里免得只躺在提交信息里）：
   ① `formatAmount` 的分组校验偏松 —— `AMOUNT_RE` 的首段是不限长的 `\d+`，
   `1234,567`、`12345,678` 这类非规范分组也被接受（结果仍会规范化，不会出错值）；要收紧就改那一处正则。
   ② **下划线只认 `<u>` 标签与行内 `style.textDecorationLine`**，纯 CSS 类画的下划线不会被
   `readInlines` 读回。当前无害（`render/html.ts` 固定用 `<u>`），但**将来若把下划线的渲染改成 class
   就会静默丢格式** —— 动那处渲染前先回来看这条。
5. **W2 验收发现的次要项**（非阻断，记在这里免得只躺在提交信息里）：
   ① `MARGIN_PRESETS` 的 key / label 由 `default` / `gov`、「四边 25mm」改成了 `manager` / `govDoc`、
   「管理人文件」—— 与 4.1 节「保住已发布的 lib 接口」只部分相符。仓库内没有消费方，且
   `package.json` 是 `private: true`（并没有真的对外发布过），所以暂不处理；将来真要发版前再定键名。
   ② `scripts/verify-browser.mjs` 里硬编码了并排间距 `18px`（数值源自 `render/css.ts` 的 `gap`）。
   ③ 切模板的断言分散在 `verify-browser` 与 `verify-editor` 两个脚本里。
   ②③ 属整洁度问题，不影响结论。

6. **页尾删除会让断点处的字符重复（分页缺陷，2026-09-13 复现并定位，未修）**。在一个跨页长段落
   （demo 样本里那段「关于对外投资部分，…」正文）的**页尾**连按 Backspace，断点处的字符会被复制进模型、
   越删越多（用户截图里是 `…被告的的民的的民…`），模型长度在 −1／+1 之间震荡。
   **根因是三条叠在一起**：① 删除后新分页里这一格的 `from/to` 常常与上一轮**完全相同**（上一页末行让出的
   字位正好被下一页首字顶上来）；② `retagFragments` 在 `input` 时**命令式改写** `data-from`／`data-to`
   （值是重排前的），而 Vue 重渲染时按「新旧 prop 值相同」**跳过属性更新**，于是属性停在旧值、`v-html`
   却按新分页重渲染 → **DOM 文字与坐标差一个字符**；③ 之后每次读回都把 919 字的 DOM 塞进 918 字的区间，
   `replaceRange` 把多出的字符复制进模型。同一毛病的另一半：`retagFragments` 的
   `data.to = cursor + prefix + text` 而 `textContent` 已含 `.wtp-num` 前缀，**前缀被算两遍** ——
   编号段落跨页时表现为**丢字**（不是重复）。
   **修的方向**：问题不在算术，而在「片段的 `data-from/to` 有两条权威」—— 让分页结果成为唯一权威
   （重渲染后无条件把属性同步回 `pages.value`，或让 `retagFragments` 不再写属性），并确认按
   `(blockId, from)` 编的 v-for key 不会让 Vue 复用错元素；验收要加一条「跨页段落页尾连按 N 次 Backspace，
   模型长度必须恰好 −N（现在会少减甚至增加），且断点窗口不得出现重复」。
   复现探针（只读、**临时件未提交**、产物在 `.qwen/tmp/`）：`probe-pagination.mjs`（逐步快照）、
   `probe-pagination2.mjs`（`beforeinput`/`input`/mutation 事件时序）、`probe-pagination3.mjs`（dump 片段
   DOM 文字，给出 919/918 的铁证）。更完整的机制说明见项目记忆 `pagination-boundary-duplication.md`。
7. **W4a-1 验收发现的三处非阻断项**（记在这里免得只躺在验收结论里）：
   ① **`columns` 只能由 body 行回推**：手搓模型里 `columns` 大于所有 body 行的格数时，md 往返会把它收窄
   （md 语法没有表达「总列数」的地方）。真实路径（手写 md、将来的 UI）不受影响；真要修就给围栏加个 `columns=` kwarg。
   ② **表格里的批注锚点**在模型 / md / docx 三层都已支持，但**样本与 Word 对账没覆盖**（为避开 `assert-docx`
   里「批注 1 条」那条既有断言）。需要时再补一条样本。
   ③ 半截 `[[`（未闭合的批注括号）会把其后的 `| b |` 一起吞进同一格 —— 内容不丢、往返稳定，
   只是格数比旧算法少一个，记录备查。
8. **W4a-2 验收发现的两处非阻断项**（非阻断，记在这里免得只躺在验收结论里）：
   ① **普通段落「整段渲染」时尾部软换行会被丢掉**：恰好落在切片右端的尾部软换行被 `sliceInlines` 丢掉，
   量测按 2 行算、实渲 1 行。表格格不走切片，不受影响 —— 属「软换行只在单元格内支持」的范围内。
   修片尾规则会破坏真跨页的记账，本波**不修**，等软换行要不要推广到普通段落时一起处理。
   ② **md 参差行补出的 `<td>` 是幻影格**：md 里后行比首行多格时，渲染会补空 `<td>`，但模型里没有这些格 ——
   在里面打字不落模型、下次重建即丢。该路径要浏览器才能验，本机验不了，留待有浏览器时确认。
   ③ `verify:p2` / `p3` / `pages` 本波**一律没跑**（Edge 坏），`verify-browser` / `verify-editor` 里改动的断言
   全部属「写了没跑」；表格版式一致性改**人工核对**（见 6.9）。

---

## 9. 本机环境坑（踩过的，照做能省很多时间）

1. **宿主 Qwen Code 会崩在 TUI 的 React 渲染循环上**。调试日志
   `C:\Users\18082\.qwen\debug\<session-id>.txt` 里有 `[STARTUP] [UNCAUGHT_EXCEPTION] React error #185`，
   栈全在 React 的提交/状态更新路径（`getRootForUpdatedFiber` → `dispatchSetState` → `commitRoot`），
   即终端 UI 层崩，**与内存无关**（32 GB 机器还剩 12 GB 时照样崩）。
   2026-09-13 一天之内崩了三次，每次都紧贴「代理回报一份极长报告」或「长输出灌进 TUI」。
   ⚠️ **当天傍晚又崩两次（W4a 会话），触发了对上面这句的订正**：那两次都落在**后台任务落地那一刻**，
   而当时子进程的 stdout 早已 `> .qwen/tmp/*.log 2>&1` 全量重定向（有一次主 session 手上一块大文本都没有），
   所以触发面比「大块输出」更宽 —— **「后台任务完成 → 自动续跑/刷新界面」这条状态更新路径本身也不稳**。
   由此多两条硬纪律：① 主 session 读取一律**有界**（只 `git status --short` / `git diff --stat` / 定向小段 /
   用 node 过滤后打印 ≤25 行；完整 diff、长日志、结论的长证据段只看文件路径），子进程报告**不整段转述**，
   **主 session 自己的回复也必须短**（它不受任何截断保护）；② 想彻底绕开那条路径，可让子进程用 `start /b`
   完全脱离（不注册成 harness 的后台任务）再由主 session 轮询结果文件，代价是失败只能靠轮询发现。
   完整的排查与可用的规避开关（`ui.shellOutputMaxLines`、`tools.truncateToolOutput*`、别按 `Ctrl+O` 等）
   记在 `C:\Users\18082\.qwen\qwen-tui-large-output-crash.md`（用户 2026-09-13 选择暂不改设置）。
   对策：让命令输出重定向到文件、代理报告限长（见第 0 节）；**长任务做完后另开新 session**，
   别让一条进程链背着几千万 token 的 transcript 继续往前推。
2. **命令环境**：Windows 11 + cmd.exe（不是 bash），路径统一用正斜杠；所有命令加 `rtk` 前缀
   （它是 token 压缩包装，没有对应过滤器时原样透传，始终安全）。
   ⚠️ `rtk` 成功时也可能打印 `Error: (none)` 之类的文本，**判断成功看 exit code（0 = 成功）**。
3. **`npm run verify` 很重**：`verify:p1` 会真的用 Word COM 打开 docx 读回格式（本机需装 Word），
   `verify:p2/p3` 会真的用系统 Edge 起无头浏览器。跑一次全量约十几分钟，别在循环里反复跑。
   注意：`verify:p1` 那套脚本必须保持纯 ASCII。
4. **prettier**：仓库 `format` 脚本用 `npx prettier --experimental-cli`。
   默认 CLI 会对**你根本没碰过**的既有文件报格式问题（`numbering.ts` / `paginate.ts` / `measure.ts` 等），
   那属既有状态，不要顺手格式化它们。
   ⚠️ **不要对既有文件跑 `prettier --write`**：根 `.prettierrc.json` 写的是 `tabWidth: 4`，而源码实际是
   **2 空格缩进**、也不满足 prettier 的若干其它规则 —— 一旦 `--write`，整批文件会被按 4 空格重排。
   （2026-09-13 W4b-1 会话踩过：9 个代码文件被重排，靠 `git -c core.autocrlf=false checkout --` 退回 HEAD、
   再用**格式化之前**的工作区快照 `.qwen/tmp/w4b1r-baseline.diff` 的对应 hunk `git apply --include=` 回去，
   才做到逐字节还原。要动格式化，先想清楚这一步。）
   另外 **prettier 会忽略 `.qwen/` 下的文件**：把 HEAD 副本拷进 `.qwen/tmp/` 再 `prettier --check`，会得到
   「全部合规」的**假阴性**，不能拿它判定某个文件是否 prettier-clean。
5. **字体依赖本机安装**：预览与 docx 都写「仿宋 / 华文中宋 / Times New Roman」，都是商业字体、不能内置。
   「方正小标宋」这类公文字体本机大概率没有 —— 涉及新字体前先确认，否则预览会回退、行宽会变。
6. **`docx` 库打包后再改 `styles.xml` 时必须 `createFolders: false`**（见 `lib/docx/lineUnits.ts`），
   否则会多出目录条目，Word 打开这种 docx 会**卡死在 `Documents.Open` 不返回**。
   （历史教训：`w:beforeLines` / `w:afterLines` 只能靠这种方式补写，库不暴露这对属性。）
7. **本机 Edge 的程序化启动从 2026-09-13 13:42 起坏了**（`msedge.exe --version` 零输出；用新建目录做
   `--headless=new --dump-dom about:blank` 也零输出；playwright 无头/有头一律报 `browser has been closed`；
   **窗口模式正常**，所以「Edge 能开」不等于能跑验收）。此时 `verify:p2` / `verify:p3` 的浏览器半场
   与 `verify:pages`（依赖 verify-browser 产出的预览页数）**都跑不了**；**用户要求不跑 playwright、只做代码级验证**。
   复检一条命令：`msedge.exe --version` 应打印版本号。详见项目记忆 `project/edge-headless-broken.md`。
   ⚠️ 与第 3 条一起看：第 3 条里「`verify:p2/p3` 会真的用系统 Edge 起无头浏览器」这句，现在跑不动。

---

## 10. 进度

| 波次 | 内容 | 状态 |
| --- | --- | --- |
| W1 | 快捷键组（ctrl+U / ctrl+shift+E / alt+4 / ctrl+P）+ 三种特殊空格；顺带把「列表段落」首行缩进改成 0 | **已完成**（2026-09-13，提交 `4ca96c8`；独立验收结论「可接受」，两处次要项见第 8 节第 4 条） |
| W2 | 样式与页边距绑定（两套模板）+ 双页并排 | **已完成**（2026-09-13，提交 `734dc48`；独立验收结论「可接受」，次要项见第 8 节第 5 条） |
| W3 | 查找替换 + 导航窗格 | **已完成**（2026-09-13，本次提交；独立验收先判「不可接受」——替换后插入符锚点缺失，修完聚焦复核「可接受」，结论见第 5.3 节） |
| W4a-1 | 表格：模型 + md 围栏语法 + docx 导出 + Word 读回对账 | **已完成**（2026-09-13，本次提交；独立验收先判「不可接受」——格内 `{红|…}` / `[[]]` 里的竖线被当列分隔符，修完聚焦复核「可接受」且做了变异测试；结论见第 6.7 节） |
| W4a-2 | 表格：预览渲染 + 量测 + 分页 + 编辑读回收口 + demo 样本放表 + 「插入表格」按钮 + 软换行 | **已完成**（2026-09-13，本次提交；机器验全绿，独立验收「可接受」并修掉一处必修缺陷；**版式一致性改人工核对**，结论见 6.9） |
| W4b-1 | 表格：结构操作（增删行列、unit/note radio、行高两档）+ 上下文工具条 + 格内 Shift+Enter + 边界护栏 | **已完成**（2026-09-13，本次提交；机器验全绿，独立验收与增量复核均判「可接受」，结论见 6.11） |
| W4b-2 | 表格：Tab / 方向键跨格导航（用户已定：Tab 停在最后一格时按了不作任何响应） | 未开工（设计见 6.10 末段） |
| W5 | 节编辑框架 | **需先出模型设计给用户过目**，未开工 |

**收尾待办**：`README.md` 的「待做」一节加一行指向本文件（`PLAN.md`）—— **已完成**（随 `281f979` 提交）。
往后的维护约定：每开一波之前先回来读一遍对应小节；每完成一波，把该波标题改成「（已完成 YYYY-MM-DD）」、
补一行结论、并在上面这张表里改状态。
