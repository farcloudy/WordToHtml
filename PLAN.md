# WordToHtml 开发文档（二次开发参考）

> `README.md` 是给**使用者**的：怎么跑起来、源码怎么写、快捷键、已知限制。
> 这份文件是给**改这个仓库的人**的：约束、已定稿的设计、文件地图、组件接口、验收脚本、还开着的问题、本机环境坑。
>
> **维护约定**：做完一件事就把「过程与结论」从这份文件里删掉 —— 需要考古时 `git log` 与提交信息里都有。
> 这里只留**仍然有效**的东西：约束、接口事实、还没做完的、还没修的。

---

## 1. 不可违背的项目约束（改任何东西前先记住）

1. **正常输入不得触发重排**。渲染只读 `viewDoc` / `pages` 两个浅响应式快照，编辑中的模型不参与渲染；
   每次 `input` 只重量内容变过的块，只有分页结果真的变了才重建 DOM 并把插入符按「块 id + 字符偏移」放回去。
   **新加的任何交互都不要破坏这条**（这是 P3 验收的核心断言：敲字后片段仍是同一个 DOM 节点）。
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
9. **先问 → 记快照 → 再改**。`pushHistory()` 必须在改模型**之前**调；模型层的谓词
   （`canMergeIntoPrevious` / `canJoinWithNext` 这类）先问过再动手 —— 否则撤销下来是空操作。
10. **默认值不落模型、不写 md**（全仓库统一约定）：`sections` 整篇全默认时整个字段不写、围栏
    `minLines=1` / `cantSplit=yes` 不写、格内显式写出的默认样式解析时归一化掉不落字段、
    `::editor` / `::section` 只写非默认键。
11. **项目里没有弹窗形态**：不可逆操作靠 `pushHistory` 的撤销兜底（删整表也不二次确认）。
    浮动面板只有一套做法：`beginPanelDrag` + `onPanelDrag` + `endPanelDrag`，打开时焦点交给第一个输入框，
    `×` 或 `Esc` 关闭并把焦点还给版面。

---

## 2. 已定稿的设计决策（勿再改回）

- **载入后文档至少有一个块**：零块输入（`content: ''`、手写空 md、`emptyDoc()`）由 `WordPaper` 载入时
  过一道 `ensureBodyBlock()`（`lib/types.ts`）补一个空白正文段落（`kind: 'body'`、`inlines: []`），
  `content` / `model` 两条入口都在那一处。**为什么**：编辑层靠页面上带 `data-block-id` 的片段把 DOM
  读回模型，零块 ⇒ 页面上一个片段都没有 ⇒ 读回的循环体一次都不跑，而 `.wtp-content` 自己就是
  `contenteditable`，字被插在它下面 ⇒ 字只进 DOM、永远回不到模型（看得见、保存与导出却是空的，还不报错）。
  `emptyDoc()` 本身仍返回零块 —— 它表达的是「没有内容的文档」，不替调用方决定该不该有一个空段落。
- **文件模板** = 样式与页边距绑成一体（`DOC_TEMPLATES`）。`manager`「管理人文件」四边 25mm、样式走
  `DEFAULT_SPEC`；`govDoc`「简易公文格式」上 37 / 下 35 / 左 28 / 右 26 mm、整套样式覆盖在
  `GOV_STYLES`（正文与各级标题三号 16pt、标题二号 22pt 方正小标宋简体、列表与页脚四号 14pt、
  行高一律固定 28.95pt = 网格行高）；`MARGIN_PRESETS` 由 `DOC_TEMPLATES` 派生。
  公文那套**每行 27 字**（GB/T 9704 写的是 28 字，原因见第 5 节），这是有意的偏差。
- **单元格 = 可寻址的「伪块」，格内**多段落**：一个格子装 `paragraphs: InlineHolder[]`（Word 的 `w:tc` 里本来就放得下多个 `w:p`），
  每段与段落同一个形状，所以 `edit/model.ts` 的 `findContainer(doc, id): { inlines: Inline[] }` 一行都不用改就吃掉了它。
  id 规则：第 0 段就是 `cellId(tableId, r, c)` → `` `${tableId}.r${r}c${c}` ``（既有 id 一字不改），第 N 段是 `` `${cellId(...)}.p${N}` ``（`cellParagraphId`）。
  段落 div 挂 `data-block-id`，于是 `edit/dom.ts` 的坐标换算一行都不用改。**已否决**给坐标系加第二层（`blockId + row + col + paragraph + offset`）。
  不变式：`paragraphs.length >= 1`（解析、`normalizeTable`、`cloneDoc` 都保证）；
  结构操作（`splitCellParagraph` / `mergeCellParagraph`）动手前还会**就地把坏形状补齐**
  （内部 `ensureCellParagraphs`）—— 手搓模型缺 `paragraphs` 时若只改一份临时兜底对象，
  函数会报「切好了」而模型一字未动。
  ⚠️ `cellId` 里嵌的是**行/列下标**，任何增删都会让其后的格子 id 整体位移 —— 操作后的插入符必须按新下标重算；段落下标同理（增删段会让 `.pN` 位移）。
- **格内换段与并段是独立的一对操作**（`edit/table.ts` 的 `splitCellParagraph` / `mergeCellParagraph`），
  不走 `splitBlock` / `mergeIntoPrevious` —— 那两条只认 `textBlock`，格内走它们会把整张表当段落切开。
  `kind` / `align` 是**格子级**的（Word 的 cell 级属性也只有垂直对齐与宽度），格内各段共用一套。
- **表格断行按行级**（单行不拆，可在行与行之间断开续页，`w:cantSplit` 语义），不是整表不拆。
- **软换行走真 `<w:br/>`**（零宽 `BreakInline`），正文段落与表格格子同一条路（都走 Shift+Enter）。
  md 写法是 `{br}`，**不是** `\n`
  —— `md/parse.ts` 的 `parseInline` 第一条分支就是「反斜杠吃掉下一个字符」，`\n` 会被吃成字面量 `n`。
- **表内文字缺省复用「列表段落」样式**（所以 `listItem` 的首行缩进是 0，别再给单元格加缩进豁免），
  但可以逐格换成任意 `BlockKind`（`TableCellModel.kind?`，缺省 `listItem` 不落字段）；
  `kind === 'body'` 导出时**不挂 `w:pStyle`**（`WT-Body` 就是 Word 的 Normal，与正文同一取舍）。
- **「列标题行」不独立表达**：`role` 只有 `unit` / `body` / `note`，要加粗居中就自己套格式
  （逐格对齐已能解决居中）。
  ⚠️ 上一段里「因此 Word 那种『跨页自动重复标题行』不做」**已被 issues/20260916-1 第 4 条推翻**：
  用户 2026-09-16 明确要这个功能，方案 A（表格级 `headerRows`）已过审并落地（W11）——
  模型是 `TableBlock.headerRows?: number`（前 N 行 = 标题行，缺省 0 不落字段，归一化夹到
  `[0, rows.length]`，读侧统一走 `types.ts` 的 `headerRowCount()`），docx 落成前 N 行的
  `w:tblHeader`，续页顶端由分页与渲染重复这几行（`PageFragment.headerTo`）。
- **重复标题行是「只读装饰」**：续页片段里重复渲出的那几行 `<tr>` 带 `wtp-tr-repeat`、
  `<tr>` 上带 `contenteditable="false"`、格内**不挂 `data-block-id` / `data-cell-id`**
  （同一格在版面里出现两份可寻址元素会让
  `retagFragments` / `syncPlain` 读重、整格刷选认错格）。
  `contenteditable="false"` 挡的是**改内容**，**不是**「插入符不会落进去」——实测选区仍可能落进
  重复行（落进去敲的字只存在于 DOM、进不了模型）。
  另有一条结构性不变式：续页片的重复区间恒取 `min(headerRows, rowFrom)`，
  **一片绝不重复它自己已经含有的行**（见 `paginate()` 主循环那条注释）——否则同一行被正本与
  重复行渲两遍，「DOM 行数 = 模型行数」当场不成立，而 docx 只写 `w:tblHeader`，预览与 Word 就分家了。
- **节**：`blocks` 保持扁平 + `DocModel.sections?: SectionSettings[]`（**下标 = 节号**，长度 = 分节符数 + 1）；
  `SectionBreakBlock` 只剩 `{ t, id }`。首节 `linkPrevious` 在解析时强制 `false`；
  `resolveSections()` 必须对 `sections` 缺失/偏短兜底（手搓模型、旧 md、旧 docx 读回）。
- **表格复选**的选中态**不进模型**（组件的交互状态），只在批量操作执行时改模型；
  只做**矩形块**，明确省略「拖到表格外自动扩成整行/整列」与「不连续块之间的合并显示」。
- **修订/导航开关存 md 用文档首行 `::editor`**，**不用 front matter**（`---` 在本项目已经是分节符语法）。
- **组件边界** = 顶栏 + 功能区 + 纸张（含批注侧栏、查找替换面板、插入表格面板、导航窗格、提示条），
  **不含**「类 md 源码」pane 与模式切换（demo 靠两个插槽挂回同一副外壳）。
- **F4 只重复「格式类」操作**（加粗/下划线/颜色/段落样式/格内样式/两组对齐/行高/表头附注行开关/
  重复标题行开关），
  不重复删除、并段、插表、插入空格、分节分页符、接受拒绝修订（会把撤销栈搞乱或本身不是格式操作）。
- **快捷键表的默认值**放在可手改的 `src/lib/edit/shortcuts.json`，它同时是 `shortcuts` prop 的默认值；
  `SHORTCUT_ACTIONS` 的顺序**即冲突时的优先序**（先到先得）。
- **字号/行距/页边距这类数值一律只写在 `spec.ts`**；按钮文案与键位可以分散，数值不行。

---

## 3. 文件地图（接入点）

> ⚠️ 行号会随改动漂移，**用前先 grep 核对**，别照着行号直接改。

| 想改什么 | 落在哪 |
| --- | --- |
| 新增/修改样式（字体字号行距段距缩进对齐编号） | `src/lib/spec.ts` 的 `BlockKind` / `BLOCK_KINDS` / `STYLE_KEYS` / `DEFAULT_SPEC.styles`。`Record<StyleKey, TextStyleSpec>` 会强制补齐，漏了编不过 |
| 页边距 / 页面尺寸 / 文件模板 | `src/lib/spec.ts` 的 `PageSpec` / `MarginPreset` / `DOC_TEMPLATES` / `DEFAULT_SPEC.page` / `resolveSpec()` |
| 工具栏按钮（UI） | `src/components/WtpEditor.vue` 的功能区四页（`.ribbon-tabs` / `.panel-{start,insert,layout,table}`）；demo 自己的东西在 `src/App.vue` |
| 编辑器侧的行为与对外 API | `src/components/WordPaper.vue`：`onKeydown`（**只挂在 `.wtp-content` 上，不是全局监听**）、`defineExpose` |
| 全局快捷键 | 项目里 `document` 上只有一处 `addEventListener`：`selectionchange`。要全局键盘得自己加监听 |
| 选区感知 | `src/lib/edit/dom.ts` 的 `currentRange` / `selectedRanges` / `placeCaret` / `placeRange` / `placeCaretAfterBreak` / `pointToOffset` / `offsetToPoint` / `fragmentOf`；回显走 `WordPaper.vue` 的 `emitSelection` → `selection-change` 事件 → `WtpEditor.vue` |
| 模型操作（纯函数层） | `src/lib/edit/model.ts`：`insertText` / `deleteRange` / `replaceRange` / `deleteSpan` / `splitBlock` / `mergeIntoPrevious` / `joinWithNext` / `setContainerKind` / `applyFormat` / `resolveRevisions` / `revisionSpanAt` / `hasRevisions` / `addComment` / `insertBreakAfter` / `removeBreak` / `cloneDoc` … |
| 表格结构操作 | `src/lib/edit/table.ts`（纯函数）：`findTable` / `findCell`（**不管 `.pN`**，整格操作） / `findCellAt`（格子 + 那一段） / `cellParagraphs` / `cellParagraphCount` / `emptyCell` / `splitCellParagraph` / `mergeCellParagraph` / `bodyRowIndexes` / `bodyInsertIndex` / `insertBodyRow` / `removeBodyRow` / `setRoleRow` / `insertColumn` / `removeColumn` / `normalizeTable` / `setMinLines` / `stepCell` / `verticalCell` / `setCellKind` / `setCellAlign` / `removeTable` |
| 节 | `src/lib/section.ts`（`resolveSections()` 等）+ `src/lib/edit/section.ts`（纯模型操作） |
| 分页与分节 | `src/lib/render/paginate.ts`：`PageFragment`（含 `rowFrom/rowTo`）/ `stepNumbering` / 主循环；布局比对 `sameLayout`（`WordPaper.vue` 内） |
| 量测 | `src/lib/render/measure.ts`：`measureDocument`（带增量缓存）、`MeasureCache`；**换规格表/换版心宽度时必须 `clearMeasureCache`**（缓存签名含版心宽度） |
| 分页预览的 HTML/CSS | `src/lib/render/html.ts`（行内标记 → HTML，预览/量测/编辑读回共用）、`src/lib/render/css.ts`（规格表 → 预览 CSS） |
| docx 导出 | `src/lib/docx/export.ts`：`groupSections` / `pageNumberParagraph` / `buildDocument` / `paragraphStyles` / `buildZip`（打包后补写 `styles.xml` 的 `beforeLines`）/ `toBlob` |
| md 语法 | `src/lib/md/parse.ts`、`src/lib/md/serialize.ts`。**`verify-docx.mjs` 会断言「模型 → md → 模型」往返一致**，加新语法必须双向可逆、且渲染成 md 要字节稳定 |
| 快捷键表 | `src/lib/edit/shortcuts.ts` + 默认值 `src/lib/edit/shortcuts.json` |
| 对外导出 | `src/lib/index.ts` |
| 撤销粒度 | `WordPaper.vue` 的 `Snapshot` / `pushHistory`（栈上限 200） |
| 验收脚本 | `scripts/`，命令映射见第 6 节 |
| 第三方库能力边界 | `node_modules/docx/dist/index.d.ts`（**只读 d.ts，不要猜 API**） |

模块划分（`src/lib/`）：

```
spec.ts            页面与样式的唯一真相源 + 文件模板 DOC_TEMPLATES + 单位换算
types.ts           文档模型（块 / 行内 / 修订 / 批注 / 表格 / 节）+ 批注锚定文字提取
section.ts         节清单（resolveSections：把 DocModel.sections 解析成逐节设置）
numbering.ts       中文序数编号与层级重置
md/parse.ts        类 md → 模型        md/serialize.ts  模型 → 类 md
edit/model.ts      编辑操作的纯函数层（切分／合并／替换／格式化／批注／修订／删除）
edit/table.ts      表格结构操作的纯函数层
edit/section.ts    节设置的纯模型操作
edit/dom.ts        预览 DOM ⇄ 模型的桥（读回 inline、插入符与选区的坐标换算）
edit/amount.ts     金额格式化（千分位 + 两位小数，纯函数，不碰 locale）
edit/search.ts     查找与替换（纯函数：字面量／正则、范围限定、跳过删除修订、替换留痕）
edit/outline.ts    导航窗格的大纲（h1/h2/h3 → 条目 + 指纹）
edit/shortcuts.ts  快捷键表（默认表 + 组合键解析 + 命中判定，纯函数）
docx/export.ts     模型 → docx（段落样式、页码域、w:ins/w:del、w:u、批注、表格、分节）
docx/lineUnits.ts  docx 库不暴露的 w:beforeLines 之类，打包后改 styles.xml 补写
render/css.ts      规格表 → 预览 CSS（含打印样式）
render/html.ts     行内标记 → HTML（预览、量测、编辑读回共用同一套结构）
render/measure.ts  DOM 实测：行数、行高、每行起始字符偏移（带增量缓存）
render/paginate.ts 纯函数分页：装箱 + 跨页按行切开 + 分节重编号
```

分页不是 CSS 断页，而是「实测行盒 → 装箱 → 在行边界切开」算出来的：行距是固定值，行高因此是确定量，
分页可以算出来而不是猜。孤行控制默认开启（页尾与页首各至少 2 行，因此少于 4 行的段落不会被拆开；
表格行带 `atomic` 绕开它）。

---

## 4. 组件接口

### 4.1 `WtpEditor`（对外组件）

对外就一个组件：`WtpEditor`（`src/components/WtpEditor.vue`，库里以 named export 导出）。
**顶栏 + 功能区 + 纸张**整副外壳都在它里面 —— 批注侧栏、查找替换面板、插入表格面板、导航窗格、提示条
也归它管。「类 md 源码」pane 与「所见即所得 / 源码」模式切换**不在**组件里（那是 demo 的事，
见下面「插槽」）。

props：

| prop | 类型 | 说明 |
| --- | --- | --- |
| `content` | `string` | 类 md 内容。**它是初始内容**：编辑过程中组件不回写（免得一个字回调一次）。留空 = 空字符串 |
| `fileName` | `string`（必填） | 顶栏那一行的文件名。导出 docx 的名字 = 它 + `.docx`（空名兜底「未命名」，已带 `.docx` 不叠） |
| `author` | `string`（必填） | 修订与批注的作者名 |
| `template` | `string` | 文件模板 key（`DOC_TEMPLATES` 里的一项）；留空 = 第一套 |
| `shortcuts` | `ShortcutOverrides` | 偏好快捷键表，只写要改的动作（见 README 的「快捷键」）；留空 = `src/lib/edit/shortcuts.json` 那份默认表 |
| `editable` | `boolean` | 打开编辑层，默认 `true`。**这一项是 W8 加的**（issue 的清单里没有）：组件总得有个办法表达「只读预览」，demo 的源码视图靠它 |

emits：

| 事件 | 载荷 | 说明 |
| --- | --- | --- |
| `save_md` | `string` | 顶栏「保存」按钮或 `ctrl+S`：回传 `toMd(getModel())`。`ctrl+S` 会 `preventDefault`，不弹浏览器自己的保存对话框 |
| `save_docx` | — | 「导出 docx」**真的触发了浏览器下载**之后发出（下载名按上面 `fileName` 的规则算） |
| `update:fileName` / `update:author` / `update:template` | `string` | 受控回写：顶栏文件名、修订作者、文件模板改动时回传。**不接也能用**，只是使用方拿不到新值 |
| `paginated` | `number` | 分页完成后的页数 |
| `toast` | `string` | 一句话提示（组件自己也画提示条，这个事件只是把它报出去） |
| `editor-flags` | `EditorFlags` | 模型里 `::editor` 解析出来的两个开关（载入 / 重建后发一次） |

两个插槽（demo 就靠它们把源码 pane 与模式切换挂回同一副外壳）：

- `#bar-extra`：顶栏里、文件名之后；
- 默认插槽：纸张左侧、导航窗格之后。

最小示例：

```vue
<script setup lang="ts">
import { ref } from 'vue'
import { WtpEditor } from 'wordtohtml'
import 'wordtohtml/dist-lib/wordtohtml.css'

const editor = ref<InstanceType<typeof WtpEditor> | null>(null)
const fileName = ref('关于××的情况说明')
const author = ref('张三')

function onSaveMd(md: string) {
  // 存盘或交给后端；要「随时拿最新 md」也可以直接调 editor.value?.toMd()
  console.log(md)
}
</script>

<template>
  <WtpEditor
    ref="editor"
    content="# 关于××的情况说明&#10;&#10;正文……"
    :file-name="fileName"
    :author="author"
    @save_md="onSaveMd"
    @save_docx="() => console.log('docx 已开始下载')"
    @update:file-name="fileName = $event"
    @update:author="author = $event"
  />
</template>
```

三个必知的点：

- **`content` 是初始内容**，编辑过程中不回写；要拿最新 md 调组件暴露的 `toMd()`（或 `getModel()`），
  `save_md` 交出去的就是同一个字符串；
- **`ctrl+S` 归组件管**：按它就等于点「保存」，并且拦住浏览器默认的「保存网页」；
- 组件自带整屏高度（`height: 100vh`）与内部滚动；要嵌在别处就在外面覆盖 `height`。

demo（`src/App.vue`）本身就是一份接法示例，另外用 URL 开关演示这几个入口（验收脚本也走它们）：
`?shortcuts=bold:ctrl+shift+b,formatAmount:ctrl+alt+4` 覆盖快捷键表、`?template=govDoc` 直接把模板当 prop
递进去、`?empty=1` 演示 `content` 留空。demo 把组件回传的值（`update:*`、`save_md` 收到的 md）与
`save_docx` 的次数攒在 `window.__wtpDemo` 上（`{ fileName, author, template, lastSaveMd, docxCount }`）——
浏览器脚本/控制台可以从那里读；真实使用方直接绑自己的状态即可，不必这么写。

组件暴露的方法（`defineExpose`）—— 前六个是「够用」的那一层，后面是把纸张组件的既有能力转发出来：

| 方法 | 说明 |
| --- | --- |
| `getModel()` / `toMd()` | 取当前模型 / 当前 md（`toMd()` = `save_md` 交出去的那个字符串） |
| `pageCount()` | 当前页数 |
| `exportDocx(): Promise<Blob>` / `downloadDocx(name?)` | 取 docx 二进制 / 直接触发浏览器下载（`name` 省略时走 `fileName` 的规则） |
| `getPaper()` | 底层的 `WordPaper` 实例 —— 开发期钩子（demo 的 `window.__wtpPaper`）与「还没转发到的能力」都走它 |
| `repaginate()` / `getSpec()` / `getMeasurements()` | 重量测分页 / 取生效规格 / 取最近一次量测值（排错用） |
| `focusBlock(blockId)` | 把某块滚到可视区中间（导航窗格点击用它） |
| `undo()` / `redo()` / `canUndo()` / `canRedo()` | 撤销重做 |
| `setEditorFlags({ trackChanges, nav })` | 把编辑器开关写回模型（见 `editor-flags`） |
| `setBlockKind(kind)` / `toggleBold()` / `toggleUnderline()` / `setColor(hex\|null)` / `formatSelectionAsAmount()` | 段落样式与行内格式（金额格式化同 `alt+4`） |
| `insertSpecialSpace(kind)` / `insertPageBreak()` / `insertSectionBreak()` / `insertTable(rows?, cols?)` | 插入类操作 |
| `getSelectionScope()` / `setSearch(q, opts?)` / `nextMatch()` / `prevMatch()` / `replaceCurrent(t)` / `replaceAll(t)` / `clearSearch()` | 查找替换 |
| `addCommentOnSelection(text)` / `addCommentAt(...)` / `replyComment(id, text)` / `removeComment(id)` / `focusComment(id)` | 批注 |
| `getCellSelection()` / `clearCellSelection()` | 整格复选的只读镜像与收起 |

### 4.2 内部组件 `WordPaper`（高级用法）

`WtpEditor` 内部用的是 `WordPaper`（`src/components/WordPaper.vue`，库里也一起导出）：
它只管**纸张**（分页预览 + 编辑层 + 批注侧栏），没有顶栏、功能区与那些浮层。
要把纸张嵌进自己的界面、自己画工具栏时才直接用它 —— 下面这套接口**属于内部组件**，
正常用法不必碰。

`defineExpose` 暴露：

| 方法 | 说明 |
| --- | --- |
| `repaginate()` | 强制重新量测与分页（一般不需要手动调） |
| `exportDocx(): Promise<Blob>` | 取 docx 二进制 |
| `downloadDocx(name?)` | 直接触发浏览器下载 |
| `getModel()` / `getSpec()` | 取当前模型／生效规格 |
| `getMeasurements()` | 取最近一次分页用到的量测值（行数、行高、段距），排错用 |
| `pageCount()` | 当前页数 |
| `setBlockKind(kind)` | 把选中的段落换成某个样式（标题／正文／抬头……） |
| `toggleBold()` / `toggleUnderline()` / `setColor(hex\|null)` | 给选中的文字加粗／加下划线／改色（`null` 恢复默认色）；两个 toggle 都是「全加粗就取消」的切换语义 |
| `formatSelectionAsAmount()` | 把选中的数字改写成千分位 + 两位小数（`alt+4` 调的就是它）；不是合法数字／未选中／跨段时返回 `false` 且不改模型 |
| `resolveRevisions(action)` | 接受／拒绝选区里的修订（`action` 为 `'accept'` / `'reject'`）；没选中文字时按插入符所在的那一串修订算，一处修订都没覆盖到则原样不动（也不记撤销） |
| `insertSpecialSpace(kind)` | 在插入符处插入特殊空格，`kind` 为 `'em'`（U+2003）／`'en'`（U+2002）／`'quarterEm'`（U+2005）；有选区时替换选区，取不到落点返回 `false`。**落点优先序：实时选区 → 实时插入符 → 最近一次记录的选区 → 最近一次记录落点** —— 键盘路径（`ctrl+alt+X/C/V`）与「插入」页那三枚按钮的焦点都在正文里，永远落在前两项上；后两项只服务「焦点真的离开正文」的入口 |
| `insertTable()` | 在当前段落之后（落点在格子里就插到那张表之后，取不到落点则追加到文末）插入一张空表，返回新表 id；默认 3 列 × 2 个 body 行、`minLines=2`、禁止跨页断行，插入后插入符落在第一个格子开头 |
| `insertTableRow(where)` / `removeTableRow()` | 在光标所在行的上方或下方（`where` 为 `'above'` / `'below'`）插入一行 / 删除光标所在行（光标不在正文行、或正文行只剩一行时空转） |
| `insertTableColumn(where)` / `removeTableColumn()` | 在光标所在列的左侧或右侧（`where` 为 `'left'` / `'right'`）插入一列 / 删除光标所在列（只剩一列时空转；unit/note 行整行一格，列操作不动它们） |
| `setTableMinLines(n)` / `setTableRoleRow(role, on)` | 行高两档（`1` / `2`＝最小一行 / 最小两行，**只管正文行**：表头行与附注行恒一行）；增删表头行（`role='unit'`）或附注行（`role='note'`），各至多一行 |
| `getCellSelection()` / `clearCellSelection()` | 整格复选的只读镜像（`{ tableId, cells }`，模型坐标；没有复选时 `null`）与「收起复选」；复选本身是组件的交互状态，不进模型（拖动刷选 / Ctrl+点击的交互见「表格编辑交互」） |
| `addCommentOnSelection(text)` | 给选中的文字加一条批注 |
| `addCommentAt(blockId, from, to, text)` | 按模型坐标加批注 |
| `replyComment(parentId, text)` / `removeComment(id)` | 回复／删除批注 |
| `undo()` / `redo()` / `canUndo()` / `canRedo()` | 撤销栈（组件自己维护，不用浏览器原生撤销） |
| `focusComment(id)` | 高亮并滚动到某条批注的正文锚点 |
| `getSelectionScope()` | 当前选区覆盖到的模型坐标区间（`SearchScope[]`）；面板打开的那一刻调用它，因为焦点一进输入框实时选区就没了 |
| `setSearch(query, opts?)` | 跑一次查找并存下会话（`opts.regex`／`opts.scope`）；非法正则只通过 `search-state` 事件回 `error`，不抛异常 |
| `nextMatch()` / `prevMatch()` | 移动「当前匹配」并把它滚到可视区中间 |
| `replaceCurrent(text)` / `replaceAll(text)` | 替换当前一处／全部；修订模式下按 Word 语义留痕（见下文） |
| `clearSearch()` | 清高亮与整个查找会话（切到只读时组件自己也会清） |
| `focusBlock(blockId)` | 把某块滚到可视区中间；可编辑时再把插入符放到该块自动编号之后（导航窗格点击用它） |
| `setEditorFlags({ trackChanges, nav })` | 把编辑器开关写回模型（缺省值删字段），见下面 `editor-flags` 事件那一节 |

`WordPaper` 自己的 props：`source`（类 md 源码）、`model`（直接给模型，优先于 source）、`spec`（规格覆盖）、`author`（修订与批注作者名）、
`editable`（打开编辑层）、`trackChanges`（修订模式）、`shortcuts`（自定义快捷键表，见 README 的「快捷键」一节）。
事件：`paginated`（分页完成后给出页数）、`selection-change`（选区变化时给出当前段落样式、加粗／下划线／颜色，
以及「选区里有没有修订」—— 工具栏用它回显）、
`toggle-track-changes`（按了 `ctrl+shift+E`，请调用方翻转自己持有的 `trackChanges`）、
`toast`（一句无效输入提示，它自己不做提示 UI，由调用方决定怎么显示）、
`editor-flags`（**模型载入 / 重建之后**给出 md 里 `::editor` 解析出来的两个开关，已补齐默认值；调用方据此设自己的状态）。
`editor-flags` 是单向的：用户在调用方那侧改开关时，要调 `setEditorFlags({ trackChanges, nav })` 把值**写回模型**
（否则切源码视图 / 序列化出来的 md 带不上这一行）。它不会因为 `setEditorFlags` 反过来再发一次事件
—— 那会形成回环，把用户刚改的值覆盖回旧的。`WtpEditor` 就是照这套接线把顶栏那两个开关管起来的。

文档含批注时，纸张右侧自动出现审阅侧栏（锚定文字 + 作者 + 时间 + 内容 + 回复），点条目会在正文里
高亮对应锚点并滚动过去；没有批注时侧栏不占位。

### 4.3 跨模块不变式（改任何一处之前先读）

**渲染与坐标**

- 渲染只读 `viewDoc` / `pages` 两个**浅响应式快照**；编辑中的模型不参与渲染，所以打字不会重建 DOM。
- 片段的 `data-from` / `data-to` 的**唯一权威是分页结果**：`retagFragments` 在 `input` 时按当前 DOM 现算，
  它写的坐标只在**下一次重排之前**有效；重排之后由 `applyFragmentRanges()` 按 `pages.value` 把每个段落片段
  无条件重写一遍（放在 `refreshLayout` 的 `nextTick` 里、`placeCaret` **之前**，没有 anchor 也要跑）。
  不这样做就会留下「Vue 按旧 vnode 跳过属性写入」的脏值 —— 那是「页尾删除把断点字符复制进模型」的根因。
  表格片段与格子不归它管（**每一段**的坐标恒为 `0..该段文字长度`，整段重渲染）。
- 判断插入符是否顶到片段左/右缘，必须按 **`fragmentOf(选区起点)`** 找片段；不能用 `fragmentAt`
  （相邻两片共用边界时会命中前一片，跨页那一退就接不了管）。
- `Shift+Enter`（正文段落与格内同一条路）的落点走 `placeCaretAfterBreak()` + `RefreshOptions.afterBreak`，
  **只有它（`insertSoftBreak`）与修订模式读回（锚点正落在一枚软换行上时）这两条路传 `afterBreak: true`**
  （软换行零宽，通用 `offsetToPoint` 一律还原到换行之前）。
- **尾随软换行在渲染时再补一枚占位 `<br>`**（`render/html.ts`）：浏览器不给尾随 `<br>` 开行盒，
  不补就比 Word 少一行，且光标停在它之后时敲的字会被插到它**之前**（打回上一行）。
  占位 `<br>` 不带 `wtp-br` 类，读回时被忽略。
  与它配套的两条：分页片段带 `tail`（这一片覆盖到本块最后一行），渲染侧靠它决定「整段不切片」
  （零宽让「整段」与「前半截」的字符区间完全一样，只有分页结果分得开）；`syncPlain` 在
  「这一片覆盖整个容器」时**整段照抄 DOM**，绕开 `replaceRange` 那条会把右端点上的零宽 inline
  留下的边界规则（那是「越敲越多」的根因）。

**跨块删除**

- 一律由**模型层**接管：`deleteSpan`（Word 语义：段落标记被删掉、首尾接起来、中间整段消失；修订模式下不并段、
  只标 `w:del`）与 `joinWithNext`（段尾 Delete = 删段落标记，与 `mergeIntoPrevious` 对称）。
  接管点在 `WordPaper.vue` 的 **beforeinput**（跨片段选区）、**compositionstart**、
  **keydown Backspace / Delete**、**onPaste** 四处，且一律 `preventDefault`。
  浏览器原生删除会把相邻片段元素并成一个、其余直接删掉，而 Vue 手里还留着那些节点的 vnode —— 页面再也补不回来。
- 每页一个 `contenteditable`，原生的删除**跨不过页边界**，页尾 Delete / 页首 Backspace 原本是空操作。
- 选区里**夹着表格格子**时退回「各容器各自删掉选中的文字」（段落的合并／切开只对 `textBlock` 有意义；
  格内的并段走 `mergeCellParagraph`，跨格的段落标记本来就删不掉）。
- 下一块是表格 / 换页标记时 `joinWithNext` 返回 `null`，调用方**不改模型也不放行**。

**表格的 DOM 契约**

- 表格片段外层只挂 `data-table-id` / `data-row-from` / `data-row-to`，**不挂 `data-block-id`**
  —— 否则 `fragmentOf()` 向上取「最近的 `data-block-id`」会把外层当成片段。
  格内是**两层**：包装层 `<div class="wtp-cell">` 是给脚本 / CSS 用的稳定钩子（**不挂 `data-block-id`**），
  里面**每段**一个 `<div class="wtp-cellpara wtp-<kind>" data-block-id data-from data-to>`；
  `<td>` 恒挂 `data-cell-id`（= 第 0 段的 id）。
- 片段渲染的 **`v-for` key 必须带 `rowFrom` / `rowTo`**（同一张表各页片段的 `from` / `to` 都是 0，会撞键）。
- **重复标题行**（续页顶端的 `w:tblHeader` 副本）另有一套规矩：`PageFragment.headerTo` 给出要重复的
  行区间 `[0, headerTo)`（分页侧恒有 `headerTo <= rowFrom`：一片绝不重复它自己含有的行），
  渲染侧**只在本片 `rowFrom > 0` 时**渲，且必须在**同一个 `<table>`** 里
  （一页一张表这条不能破）；这些 `<tr>` 带 `wtp-tr-repeat`、`contenteditable="false"`，
  **格内不挂 `data-block-id`、`<td>` 不挂 `data-cell-id`**（见上面「只读装饰」那条），
  唯一的例外是 `sameLayout` 必须一起比 `headerTo`（漏比会在「改 headerRows 却不重建 DOM」时留下旧版面）。
- 段落的 `data-block-id` 是 `cellId(...)`（第 0 段）/ `cellId(...).pN`（第 N 段）。
  包装层**刻意没有任何 CSS 规则**：不给 margin / padding / min-height，段距与缩进全靠每段那条
  `wtp-<kind>`（Word 在格内也逐段算段前段后），空段靠 `renderInlinesHtml` 的占位 `<br>` 撑高。
  `<td>` 上是 `wtp-td` + `wtp-td-<kind>`（表头/附注行另有 `wtp-td-plain`）。
- **行高最小值落在 `<td>` 的 `height`**（表格格的 `height` 语义就是最小高度，内容更高照样撑开，
  与 docx 侧 `w:trHeight @ATLEAST` 同义）。**不能**写在格内那层 div 的 `min-height` 上：那样 div 自己就被撑满、
  `<td>` 上的 `vertical-align` 挪的是「已经填满格子的 div」，居中／底端一点视觉效果都没有。
- **表头行与附注行的下限恒为「一行」**（`minLines=2` 也不变）—— 预览靠 `-min2` 的规则用
  `:not(.wtp-td-plain)` 把它们排除在外，导出侧同一处规则。
- 逐格对齐：**水平落格内各段的行内 `text-align`，垂直落 `<td>` 的行内 `vertical-align`**（写在段落 div 上无效）。
  缺字段时的默认值：水平 unit → right、note → left、body → 该格样式；垂直一律 top。
  `EditorSelection.table` 的 `alignH` / `alignV` 是**已解析默认后**的实际值（按钮 active 态吃它）。
  对齐不改行高、不改换行点，所以**不改页数**；格内换样式会改行高，**页数可以变**（预期行为）。
- 单元格左右内边距 **108 缇 = 5.4pt**，与 `docx/export.ts` 的 `CELL_MARGIN_TWIPS` 同源。
- 整格高亮走「**数据侧恒定、样式命令式**」：`<td>` 恒挂 `data-cell-id`，选中时由组件在命中的 `<td>` 上
  挂 / 摘静态类 `wtp-cellsel`（只画底色，不碰 padding / 高度 / 边框）。
  **不要**把选中态放进响应式 —— 容器类是 Vue 渲染的片段元素，选中态一进响应式就会重渲片段并重写
  `innerHTML`（格内 DOM 重建、插入符丢）。
- 刷选的**起点认 `td[data-cell-id]`，不是格内那层 div**：格高 33px 而格内 div 只有 16px，瞄准格子中心按下时
  `event.target` 就是 `<td>`。
- 结构操作（增删行列、表头／附注行、改样式、改对齐、删整表）在 `refreshLayout` **之后**都要补一次
  `void nextTick(emitSelection)`：点按钮会把焦点从正文拿走，`selectionchange` 未必再触发，
  不补这一下工具条的回显会停在旧值（用户看着就是「点了没反应」）。
- 一次批量（多格复选）**只记一步撤销**；增删行列与表头／附注行开关会**清空复选**（行列下标整体位移）。
- `edit/table.ts` 的纯函数是**唯一实现**，组件层不要再自己写一份。
  插入行必须先过 `bodyInsertIndex(table, at)` 把「想插在第几行」夹进数据行区间 —— 不夹的话，
  光标停在表头行时点「上方插入行」会把数据行插到表头之前，破坏 unit → body… → note 的显示顺序。
  锚点行 = `at <= 光标行 ? 光标行 + 1 : 光标行`。
- `removeBreak` 只吃 `pageBreak` / `sectionBreak`（它以前拿表格 id 也能删整表，是陷阱）；删块（含整表）走 `removeTable`。

**分节与打印**

- 量测**按节分段**（遇分节符先 flush 完 `getClientRects` 再改探针宽度）；量测缓存签名**含版心宽度**。
- `sameLayout` 必须把**逐节几何**并入比较，否则「单行节 / 空节切方向」时量测与片段完全相同 → 判「没变化」
  → 不重建 DOM（模型已横、纸还竖）。
- 打印只给**横排**节挂命名页 `page: wtp-landscape`，纵排走默认 `@page`（两者尺寸等价）。
  给纵排纸也挂命名页会**多打一张空白纸**（命名页切换会强制断页）。`@page wtp-portrait` 这条规则留着但没人挂它。
- `PageLayout.showPageNumber` 由 `resolveSections()` 解析后原样搬运。

**docx 库（v9.7.1）的实测坑**

- 表格 `width` 的 `type` **必须显式写 `DXA`**（库默认是 `AUTO`）；`w:tblW` 只在 `options.width` 存在时才写出。
- `cantSplit` **只在 `TableRow` 上有**；行高用 `{ value: ptToTwips(...), rule: HeightRule.ATLEAST }`。
- **单元格没有水平对齐字段**：右/左/居中只能落在格内 `Paragraph.alignment`（**逐段各写一份**，因为样式与对齐是格子级的，格内每段都挂同一个样式）；垂直对齐用
  `TableCell.verticalAlign`（→ `w:vAlign`，**一格一枚**）。
- **格内多段落 = 同一个 `TableCell` 里放 N 个 `Paragraph`**（`docx/export.ts` 的 `cellParagraph` 返回数组）。
  一个 `w:tc` 里必须**至少有一个 `w:p`**，所以坏输入（格子缺 `paragraphs`）由 `cellParagraphs` 兜底成一段空段。
- `createPageSize` 在 `orientation=landscape` 时**自己互换** w/h —— 要传**纵向尺寸 + orientation**，
  写互换值会被换两遍。
- 库对**每一节**都无条件写一个空 `<w:pgNumType/>`（API 抑制不了，语义等于不写），
  所以「是否重排」的字节判据只能看 `w:start="1"` 在不在。
- 「关联前节」= **不写** `<w:footerReference>`（不传 `footers` 就真的不写，能表达继承）；
  `link=off + numbers=off` 要写**空页脚**（否则会继承上一节的页码，「无页码」做不到）。
- 打包后改 `styles.xml`（补 `w:beforeLines` / `w:afterLines`）**必须 `createFolders: false`**，
  否则会多出目录条目，Word 打开这种 docx 会**卡死在 `Documents.Open` 不返回**。见 `docx/lineUnits.ts`。
- **相邻两个 `table` 块**（中间没有任何段落）导出成相邻的两个 `<w:tbl>`，Word 会当成**一张表**
  —— 目前没修，见第 8 节。

**快捷键**

- `parseCombo` 认**命名主键**（`Left` / `arrowleft` / `←`、`Right` / `→`、`Home`）。
- 修饰键**逐项严格比对**：`ctrl+shift+B` **不再**等同 `ctrl+B`。`ctrl+shift+Z`（重做的兄弟键）
  塞不进「一个动作一个组合键」的表，留在 `onKeydown` 里兜底，表里命中优先。
- 未知动作名忽略 + 一句 `console.warn`，**绝不悄悄改默认表**；两个动作绑同一个组合键时按
  `SHORTCUT_ACTIONS` 的顺序先到先得，后到的解绑并 warn。
- `shortcuts.ts` 刻意**逐个键**从 `shortcuts.json` 取值：拼错动作名或漏键都会在 `vue-tsc` 阶段报错，
  不会退化成「少绑一个动作」。⚠️ **加载期护栏不能写在模块顶层** —— 它跑在 `const KEY_RE` 初始化之前、
  踩暂时性死区，打包后一 `import` 就抛 `Cannot read properties of undefined (reading 'test')`。
  `build:lib` 成功**不代表产物能用**，只有真 import 一次才发现（`verify:lib` 就是干这个的）。
- `repeatable(run)` 返回 boolean，失败反馈一律与按钮同源（一条 toast、模型一字不改）。
  判据要写成「**真的没有目标格/没有覆盖到的修订**」，**不能**拿 `setter(...) === false` 当判据
  —— 那个 false 还包含「本来就是同一个对齐值、又没有覆盖可清」的空转。
- `alt+4` 的金额格式化刻意**不用 `toLocaleString`**（它按运行环境 locale 选分组符与小数点，
  同一段文字在不同机器上会输出不同结果）。

---

## 5. 样式规格表与 docx 映射

规格表在 [`src/lib/spec.ts`](src/lib/spec.ts)，是预览 CSS 与 docx 样式的**唯一来源**，
`resolveSpec(override)` 做逐层合并。两套文件模板各一份样式值：

**管理人文件**（`DEFAULT_SPEC`）：

| 类型 | 字体（中文／西文） | 字号 | 加粗 | 对齐 | 首行缩进 | 行距 | 段前／段后 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 文本标题 | 华文中宋／Times New Roman | 18pt | ✓ | 居中 | — | 26pt 固定值 | 1.5 行／1.5 行 |
| 一~三级标题 | 仿宋／Times New Roman | 14pt | ✓ | 两端 | 2 字符 | 25pt 固定值 | 0.5 行／0.5 行 |
| 正文 | 仿宋／Times New Roman | 14pt | — | 两端 | 2 字符 | 25pt 固定值 | 0.5 行／0.5 行 |
| 抬头 | 仿宋／Times New Roman | 14pt | — | 左 | — | 25pt 固定值 | 0.5 行／0.5 行 |
| 落款 | 仿宋／Times New Roman | 14pt | — | 右 | — | 25pt 固定值 | 0.5 行／0.5 行 |
| 附件 | 黑体／黑体 | 14pt | — | 两端 | — | 25pt 固定值 | 0／1 行 |
| 列表标题 | 仿宋／Times New Roman | 10.5pt | ✓ | 居中 | — | 12pt 最小值 | 0／0 |
| 列表段落 | 仿宋／Times New Roman | 10.5pt | — | 两端 | — | 12pt 最小值 | 0／0 |
| 页脚 | 宋体／Times New Roman | 9pt | — | 居中 | — | 单倍 | 0／0 |

**简易公文格式**（`DOC_TEMPLATES[1].spec.styles`，即 `GOV_STYLES`；字号取自 issues/20260916.md）：

| 类型 | 字体（中文／西文） | 字号 | 加粗 | 对齐 | 首行缩进 | 行距 | 段前／段后 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 文本标题 | 方正小标宋简体／Times New Roman | 22pt（二号） | — | 居中 | — | 28.95pt 固定值 | 0／1 行 |
| 一级标题 | 黑体／Times New Roman | 16pt（三号） | — | 两端 | 2 字符 | 28.95pt 固定值 | 0／0 |
| 二级标题 | 楷体／Times New Roman | 16pt | — | 两端 | 2 字符 | 28.95pt 固定值 | 0／0 |
| 三级标题 | 仿宋／Times New Roman | 16pt | — | 两端 | 2 字符 | 28.95pt 固定值 | 0／0 |
| 正文 | 仿宋／Times New Roman | 16pt | — | 两端 | 2 字符 | 28.95pt 固定值 | 0／0 |
| 抬头 | 仿宋／Times New Roman | 16pt | — | 左 | — | 28.95pt 固定值 | 0／0 |
| 落款 | 仿宋／Times New Roman | 16pt | — | 右 | — | 28.95pt 固定值 | 0／0 |
| 附件 | 黑体／黑体 | 16pt | — | 两端 | — | 28.95pt 固定值 | 0／1 行 |
| 列表标题 | 仿宋／Times New Roman | 14pt（四号） | ✓ | 居中 | — | 16pt 最小值 | 0／0 |
| 列表段落 | 仿宋／Times New Roman | 14pt | — | 两端 | — | 16pt 最小值 | 0／0 |
| 页脚 | 宋体／Times New Roman | 14pt | — | 居中 | — | 单倍 | 0／0 |

公文那套特有的三件事：

- **行高 28.95pt 就是文档网格行高**（`GOV_LINE_PT` = `page.gridLinePt` = 579 缇）：上 37 / 下 35 的
  版心高 637.795pt ÷ 22 = 28.99pt，**取 28.95 才装得下 22 行**（22 × 28.95 = 636.90pt），取 29pt 只剩
  21 行。Word 里「单倍行距 + 文档网格每页 22 行」与此等价，但浏览器没有网格吸附，只能写成固定值 ——
  这条等价关系由 `verify:pages` 拿预览页数与 Word 页数逐套模板对账（现在是 4/4 与 5/5）。
- **每行 27 字，不是 GB/T 9704 的 28 字**：版心宽 156mm = 442.20pt，三号字每字 16pt，28 字要 448.00pt。
  Word 能做到是因为它用文档网格**压缩字符间距**，浏览器的排版引擎没有这个能力。左右边距各让 2mm
  （各 25mm → 版心 160mm）就能精确容下 28 字，但那样版心就不是公文标准的 156×225mm 了 ——
  2026-09-16 用户选择保标准版心。顺带一条实测：这些中文字体的汉字宽度都是整 1em（仿宋/黑体/楷体/
  方正小标宋都一样），所以「每行几个字」只取决于版心宽 ÷ 字号，与选了哪个中文字体无关。
- **列表那两条的「最小值 16pt」是 14pt 仿宋的单倍自然行高**（1.1406 × 14 = 15.97 ≈ 16；管理人文件那条
  12pt 对 10.5pt 同理：1.1406 × 10.5 = 11.98）。表格格默认用 `listItem`，所以「最小一行」= 16pt、
  「最小两行」= 32pt。取整到 16pt 还有个好处：Word 的 `atLeast 16pt` 与预览的定值 16pt 恰好相等，
  两侧行高不差（若取 16.8pt，Word 那边仍按自然行高 15.97pt 排，预览会比 Word 高 0.8pt）。

- 页面默认 A4，四边 25mm（= 「管理人文件」模板的页边距）。文件模板见第 2 节：`DOC_TEMPLATES` 把样式与
  页边距绑成一体，`MARGIN_PRESETS` 由它派生（四边各自取值）；`page.margin` 也可以直接覆盖。
  页码段落用内置「页脚」样式（管理人文件居中 9pt / 简易公文格式居中 14pt 四号宋体）。
- 样式名对照：

| 块 | docx 样式名 | 来源 |
| --- | --- | --- |
| `title` | 标题 | Word 内置（Title） |
| `h1` / `h2` / `h3` | 标题 1 / 标题 2 / 标题 3 | Word 内置（Heading 1-3），自动带上大纲级别 |
| `body` | 正文 | Word 内置（Normal）。正文段落不挂样式，格式定义在默认样式上 |
| `listItem` | 列表段落 | Word 内置（List Paragraph） |
| `footer` | 页脚 | Word 内置（Footer） |
| `salutation` / `signature` / `listTitle` | 抬头 / 落款 / 列表标题 | 自定义（Word 无对应内置） |

- **「内置」的判定依据是 `w:name` 与 Word 的本地化名逐字相等**（含「标题 1」中间那个空格），
  不是 `w:styleId`。所以样式 id 仍保留 `WT-` 前缀 —— 既不影响内置归属，又能避开 `docx` 库的坑：
  styleId 一旦撞上它的内置样式表（`Heading1`、`Title`），它会额外注入一份自己的默认定义，
  同一份 `styles.xml` 里就会出现两个同 id 的 `w:style`。**反向解析按样式名匹配。**
- 命中内置样式后 Word 会自动补上大纲级别，标题因此能出现在导航窗格里，这是改内置名顺带拿到的。
- **自动编号是写在段首的真实文字**（不是 Word 原生多级列表，见第 8 节）：一级 `chineseDot`「一、」、
  二级 `parenChinese`「（一）」、三级两种 —— `arabicDot`「1、」（管理人文件）与
  `arabicPeriodSpace`「1. 」（简易公文格式，GB/T 9704 的写法，**句点后带一个空格**）。
  反向导入剥段首编号时两种写法都认（`numbering.ts` 的 `H3_NUM_RE`），否则重新编号会出现「1. 1. 」。

行内标记的落点（`md → 模型 → docx`）：`**粗**` → 加粗运行、`__下划线__` → `TextInline.underline`
→ `<w:u w:val="single"/>`、`{+…}` → `w:ins`、`{-…}` → `w:del`、`{br}` → `<w:br/>`（Word 读回来是 `chr(11)`）、
`[[文字|批注]]` → 批注锚点。**格内另有段落标记 `{p}`**（只在表格格里认，正文里不是语法）
→ 同一个 `w:tc` 里的下一个 `w:p`。`render/html.ts` 是「预览 / 量测 / 编辑读回」三处共用的同一套结构，
所以改 HTML 结构等于同时动这三处。

---

## 6. 验收（命令与断言口径）

| 命令 | 验的是什么 |
| --- | --- |
| `npm run verify:p1` | 生成 docx → **Word COM 打开** → 逐项读回 Word 实际生效的字体／字号／行距规则／段距／首行缩进字符数／样式归属（含 Word 是否把它认成内置样式）／页码域／修订／批注／**逐段下划线**（`w:u` 是否真的生效）／**表格**（行数、每行格数与整行合并、`cantSplit`、行高规则与**逐行磅值**、格内边框线型、**逐格样式与两组对齐**、表格总宽 ≈ 版心宽），与规格表对账；另在 OOXML 层面核对 `<w:u w:val="single"/>` 的处数、**软换行 `<w:br/>` 的处数**（Word 读回来的格文字里应是 chr(11)）与 `<w:tblW w:type="dxa">`／`w:tblLayout`／`w:trHeight`（逐行 = (**正文行** ? `minLines` : 1) × 该行各格样式行高的最大值 —— 表头行 / 附注行恒一行）／`w:gridSpan`／`w:pStyle`／`w:jc`／`w:vAlign` 与模型一致（**`w:pStyle` 与 `w:jc` 逐段计数：格内多段落时每段各一份，`w:vAlign` 是格子级、一格一枚**）；**格内多段落**另单独造一份最小 docx 在字节层验（同一个 `w:tc` 里 2 个 `w:p`、逐段 pStyle / `w:jc`、每格恰好一枚 `w:vAlign`、两段的文字各成一段；⚠️ 样本表里没有多段落格，所以 **Word COM 那一侧没对账**，见第 8 节）；**逐节**（W5）：Word 读回每节的纸张方向与页脚文本、`<w:pgSz>` 里 `w:orient` 与横排时宽高的互换、`<w:pgNumType w:start="1">` 只在声明「从 1 开始」的节出现、`<w:footerReference>` 只在独立设页脚的节出现（关联前节**不写**才继承），期望值一律由 `resolveSections()` 现推 |
| `npm run verify:p2` | 分页器单测（含**表格行按行装箱、原子项不吃孤行控制、同页同表行合并、跨页断开**）+ 真实浏览器（Chrome，可用 `WTP_BROWSER` 覆盖）实测：**两套文件模板各自**的样式与版心宽（`contentBoxPx`）对账、每页不得溢出、页首与续排的间距豁免、**量测值与渲染值逐块对账（含表格：各行量测高合计 = 渲染出来的表格高）**、切模板后**版心几何必须换一套**（页数可以巧合地相同，不作判据）、宽视口（2200px）下**真的有两页同处一行**（同 `top`、间距 18px、纸宽未被压缩、页带不横向溢出）、窄视口（1000px）回落成一页一排、批注侧栏与正文锚点同源（含点击高亮）、**逐节几何（W5）**：页码 1 的出现次数 = 首节 + 声明了「从 1 开始」的节数、**命名 `@page` 只挂在横排纸上**（纵排走默认 `@page`，尺寸等价；挂上 `wtp-portrait` 会让每份纵排文档多打一张空白纸）且打印媒体下页数与 `break-before` 不被改坏、用「节」工具条把末节改成横排后**只有那一节的纸宽高对调**且纸数不变、横排纸不被压窄、还原 |
| `npm run verify:p3` | 编辑操作纯函数单测（含**软换行的切片／往返、容器泛化（格内读改）、跨格查找替换**，**格内多段落**（`{p}` 的 md 往返与首段空 / 尾随空段边界、`cellParagraphId` / `parseCellId` 的 `.pN` 解析、`splitCellParagraph` 的切点与越界、`mergeCellParagraph` 的合并点与「第 0 段之前不并」、`findCellAt`、`allInlineHolders` 与 `cloneDoc` 都走遍格内每一段、`normalizeTable` 补齐坏格子、格内每一段的 pStyle / w:jc 渲染计数），以及 **`stepCell`/`verticalCell` 的跨格步进（含幻影格跳过）、`removeTable`、`removeBreak` 收紧、`setCellKind`/`setCellAlign` 与格首指令 `{@…}` 的 md 往返**）+ 真实浏览器实测：**敲字后片段仍是同一个 DOM 节点**（证明正常输入没有重排）、插入符落在刚敲完的字后面、行数变了要重排时插入符按坐标找回、回车分段／退格合并、加粗／下划线／改色、下划线的 md 往返、修订模式的增删标记、**接受／拒绝修订**（按钮只在选区含修订时亮、光标贴在修订串末尾也算、拒绝删除修订、接受插入修订、别处修订一个都没动、撤销能还原）、加批注（模型／侧栏／正文锚点三者对得上）、撤销重做、`ctrl+U`／`ctrl+shift+E`／`alt+4`（含无效输入的提示条）、三种特殊空格的码点（单测另外确认它们原样写进 docx 的 `document.xml`）、**切文件模板**（版心几何与量测行数都跟着换、插入符与选区按坐标找回、文字没丢）、打印（隐藏外壳**含查找面板、插入表格面板、表格上下文工具条与导航窗格**、真打一份 PDF 数页数并与版面页数对齐）、以及编辑后每页仍不溢出；**查找替换**（ctrl+F 开面板并聚焦查找框、匹配有高亮且当前那一处单独一层、查找不重建版面 DOM、上一个／下一个移动当前匹配、非法正则只弹提示且不高亮也不改模型、范围「当前选中的文本」只命中选区内、ctrl+G 聚焦替换框与替换一处、全部替换后模型与版面逐块一致）；**导航窗格**（条目数/编号前缀/文字/层级与模型算出的逐条一致、窗格 200px、点击跳转后插入符落在被点块的自动编号之后、正文打字不重建左栏、折叠按钮与顶栏开关）；**表格**（渲成 `.wtp-tableFrag`、外层不挂 `data-block-id`、各片渲染行数 = 行区间、**格内每一段各挂一个 `data-block-id` 而包装层 `.wtp-cell` 一个都不挂**、格内打字同步到模型且不重排、**插入表格面板**：默认 2 行 3 列、`Esc` 只关面板不插表、选 4 行 2 列后新表按规格生成且插入符落在第一格、越界值两侧边界各测一次（999 行 / 0 列 → 30 行 / 1 列，0 行 / 999 列 → 1 行 / 12 列）；**表格编辑交互**：表格页常驻（光标不在格内时给提示、按钮全灰；进格后提示收起且按钮可用）、落点提示「第 N 行第 M 列」、点「下方插入行」后模型行数 +1 且 DOM 多一 `<tr>`、插入符仍在原格内容里、删列到最后一列时按钮禁用、行高与表头／附注 radio 各开关一次并把 `-min1`/`-min2` 反映到 DOM 类名、格内 Shift+Enter 后敲字进入新行（模型里排在 `break` 之后）、**正文段落里的 Shift+Enter**（不切段、末尾多一枚软换行、版面真的占两个行盒、换行后敲的字排在 `break` 之后且不把软换行复制成两枚）、格首 Backspace 与格尾 Delete 后模型与 `<td>` 数都不变；**格内多段落（W10）**：格内 Enter 在落点把该段切成两段（块数与表格行数都不变、隔壁格不被牵连、插入符落到 `.p1` 且那时它还空着、DOM 里多出一个 `.wtp-cellpara.p1`）、在新段里打字只落在新段、格内三段时**量测行高跟着长**（行高按格内各段总高算）、回车与输入**各记一步撤销**（四级梯子逐级回退到 1 段）、第 1 段的段首 Backspace 与第 0 段的段尾 Delete 都**并成一段**且插入符停在合并点（`<td>` 数不变）、`ctrl+Z` 能把并段退回来、最后一段的段尾 Delete 仍是硬护栏、在第 1 段里按 Tab 照样跨到下一格而 Shift+Tab 落到**上一格最后一段的段尾**；**表格收尾（W4b-2）**：Tab 行优先跨格且**最后一格按 Tab 后选区与模型都不变**、Shift+Tab 回上一格末尾、格首 ← / 格尾 → 跨格而格内中间不跨、点「删除表格」后模型少一个 `table` 块且 DOM 里不再有 `.wtp-tableFrag`、插入符落在上一块末尾、格内点样式 chip 后**只有那一格**的类名从 `wtp-listItem` 变成 `wtp-h2`、两组对齐按钮的 active 态与写进 DOM 的行内 `text-align` / `vertical-align`（含「再点同一个值 = 清覆盖」与「只作用于光标那一格」）；打印隐藏里也逐个验了新增的「水平 / 垂直 / 删除表格」按钮确实不显示）；**「节」工具条（W5）**：编辑模式常驻、节号回显「第 N 节 / 共 M 节」、首节「关联前节」与「从 1 开始」整组置灰、样本第 2 节（`--- link=off restart=on`）回显「否 / 是」、点「横向」后只有那一节的纸宽高对调且纸数不变、点「页码=关」后该节不再有页码元素而别的节照旧、来回切完还原且**默认值不落模型字段**、文末插分节符的空白页页码接着前一节往下数（新分节符默认不重排）、两档命名 `@page` 的 `size` 恰好互换且外边距归 0）；**删除的边界（issues/20260915）**：跨页段落**页尾连按 5 次 Backspace** 后模型长度恰好 −5、断点窗口不出现重复，且每一步都逐块核对「版面文字 = 模型文字」（坐标与 DOM 文字不能再差一个字符）；**页尾 Delete** 删掉的是下一页片段的首字（模型 −1，不是空操作）、**页首 Backspace** 删掉的是上一页片段的末字；**段尾 Delete** 把下一段接上来（块数 −1、文字拼接正确）、**段首 Backspace** 合并前一段；**跨段选区删除**后首块 = 「首块切点之前 + 末块切点之后」且中间各块消失；**全选整页删除**后该页的块合成一段（文字清空）且后面的文字前移；**粘贴替换跨块选区**后原选区被粘贴内容取代（不再只是插在选区开头）；上述每个动作再各 `ctrl+Z` 一步，验模型能原样退回（撤销快照必须记在改模型之前 —— 并段类操作一进去就动模型，记晚了撤销就是空操作）；**功能区标签页**（四枚标签「开始／插入／布局／表格」、默认停「开始」、四页常驻且同一时刻只有一页在 DOM 里、各页各管一摊、光标进出表格时版面不上移也不下移、打印时四页外壳逐页切过去验真的都不显示）；**格内垂直对齐**（最小两行 + 单行文字时三档对齐的文字顶端偏移各自落在 0 /（格高−文字高）/ 2 / 格高−文字高 上，且三档格高一致 —— 旧写法三档偏移完全相同）；**W6**：功能区四页（开始／插入／布局／表格）逐页量 `offsetHeight` **相等**（±1px）且都没被挤成两行、自定义快捷键表（URL 参数把表递进 demo；改绑后新组合生效、旧组合失效、未知动作名不改默认表）、`F4` 重复上一步（无可重复操作时空转 + 提示条且模型一字不改、换一处选中后重放也加粗、`ctrl+Z` 只退掉这次重放）、顶栏文件名（不再有 slogan、名字可编辑、导出名 = 名字 + `.docx`、空名兜底 `未命名.docx`、写了后缀不叠）、`::editor` 与顶栏两个开关的双向同步（源码里 `trackChanges=on` → 顶栏勾上且模型里为 true；取消勾选后序列化里那一行消失；`nav=off` → 导航窗格收起 + 序列化里出现 `nav=off`）；**W7 表格复选多格**：从格内拖到对角格**刷出 2×2**（模型坐标下恰好 4 格、DOM 上恰好 4 个 `<td>` 带高亮类、两者按 `data-cell-id` 对得上）、拖出格边界后**没有残留的原生选区**、刷选前后**每一页的 `offsetHeight` 与逐块量测值都不变**且版面片段仍是原来那些 DOM 节点（= 刷选不重排、不重建 DOM）、**在同格里拖动仍是普通的选文字**（原生选区真的选中了字、选中格数为 0）、`Ctrl+点击`追加「锚格↔点击格」的矩形（9 格）与**再点已选中的格 = 去掉那一块**、`Esc` 与点正文别处都能清空（高亮同时清干净）、**批量「居中」只改那 4 格**（隔壁格一个都没动、DOM 行内 `text-align` 逐格对得上、按钮 active 态跟上、复选保留、页数不变、`ctrl+Z` 一步退回整批）、**批量换样式只改那 4 格**（类名恰好 4 个、一步撤销退回）、**跨页表**（同一张表两片上都画到高亮）、**表头行 / 附注行在 `minLines=2` 时格高仍是一行**（正文行才是两行；切成最小一行后三种行都一样高，期望值由 `ptToPx(linePt)` 现推）；**W8 组件打包**：`template` prop 决定版心几何（与规格表现推的 `contentBoxPx` 对齐）与量测行数、`fileName` prop 决定顶栏与导出名、`author` prop 决定新加批注的作者名、`update:fileName` / `update:author` / `update:template` 把改动回传给使用方、`save_md` 收到的字符串 = `toMd(getModel())`（开修订模式时带 `::editor trackChanges=on`）、`ctrl+S` 与「保存」按钮一致且 `preventDefault` 被探针读到、`save_docx` 事件发出、`content` 留空（`?empty=1`）不报错且版面出得来；**W9 新增的二十个 `ctrl+alt+…` 组合键**：段落样式 `1/8/0` 写进模型的 `kind`、标红与取消颜色（`FF0000` / 颜色字段消失）、三种特殊空格的码点 `2003/2002/2005`、分节符与分页符都落在插入符所在段落之后、格内两组对齐 `H/J/K` 与 `←/Home/→` 写进格子的 `align.h` / `align.v`、光标不在格内与没有修订时都一字不改模型并弹出对应提示条 |
| `npm run verify:lib` | 库产物（`dist-lib/`）：在 node 里真 `import` 一次不抛异常、`WtpEditor`（与内部组件 `WordPaper`）确实导出、原有纯函数导出还在、**样式产物存在且非空**（`wordtohtml.css`，里面有外壳与 `.wtp-*` 规则）、`vue` 保持 external（产物里 `import "vue"` 而不是内联一份运行时） |
| `npm run verify:pages` | 同一份源码**逐套模板**各生成一份 docx 与 Word 比页数、节数与页码重排次数（期望值由模型推：页码 1 恰好出现「首节 1 次 + 声明了从 1 开始的节数」次；预览页数与 Word 逐套对齐），并确认两套模板的**版心几何**确实不同（切模板换整套版心的证据 —— 页数可以巧合地相同，不能拿它当判据）；默认模板之外的模板再用内置样本过一次 Word 逐项对账（`assert-docx --template`） |
| `npm run verify` | 以上全部 + 类型检查 |

`verify:p1` 会在机器上启动 Word（只读打开、读完即退出）。若本机没装 Word，这一步会失败，
其余步骤不受影响。

---

## 7. 待做

### 7.1 P4 —— docx 反向导入（唯一没做的阶段）

解压 → **按样式名反查 `BlockKind`**（内置名与自定义名同一张表，见第 5 节）→ 读 `w:ins` / `w:del`
与批注锚点 → 还原模型（`stripAutoNumber` 已备好，用于剥掉段首已有的编号避免重复编号）。
**注意正文**：导出时它不写 `w:pStyle`，导入时要把它当 `body`，不能当异常跳过。

### 7.2 编辑器后续可补的（都不阻塞当前使用）

- **回车与格式工具栏也认选区**：现在 `Enter` 落在跨块选区上时只在插入符处分段，不替换选区
  （删除与粘贴已经认了）。
- **列表续行**：回车时若当前是列表段落，新段继续列表样式（现在只有正文／列表段落会继承块型）。
- **段落级操作**：上移／下移段落、`Tab` 改缩进。
- **与浏览器原生撤销打通**，或把撤销栈按「一段连续输入」合并成一步。

### 7.3 真要对外发版前必须补的

- `package.json` 仍是 `private: true` 且**没有 `files` 字段**；使用方要按包路径引 CSS
  （`import 'wordtohtml/dist-lib/wordtohtml.css'`），发版前得把这两件事定下来。
- `MARGIN_PRESETS` 的 key / label 已从 `default` / `gov`（「四边 25mm」）改成 `manager` / `govDoc`
  （「管理人文件」）—— 仓库内没有消费方，真发版前要认下这个改名。
- `editable` prop 与两个插槽（`#bar-extra` / 默认插槽）属**公开 API 扩面**：必要性成立
  （demo 的源码视图要把版面切成只读；插槽替代法会把功能区挤进右栏），但一旦对外就等于承诺了它们的语义。

---

## 8. 遗留未决问题（仍开着）

> 每一条都是「已知、未修、且不阻塞当前使用」。**相关断言保持红着，不要改成放水版。**

1. **段间距「相加还是取大」尚未钉死**。预览与分页把相邻两段的间距算作「段后 + 段前」，
   但 2026-09-13 在真实 docx 上用 Word COM 量首行纵向位置差，5 组结果都指向 Word 其实**取两者较大者**。
   若成立，预览的段间距会比 Word 大一倍。**证据链不闭环**（用直接格式做对照时 Word 没吃写进 `<w:pPr>` 的
   `w:spacing`；从零建文档的对照也不成立），要动这一块之前先把它钉死。
   另见项目记忆 `project/word-paragraph-spacing-open-question.md`。
2. **普通段落的逐段对齐仍做不到**（表格侧已解决）。段落对齐是**样式级**的（`TextStyleSpec.align`），
   `TextBlock` 没有对齐字段。用户 2026-09-13 拍板「本波只做表格」；将来要做就照逐格对齐那套来
   （模型加可选字段 + 预览行内 `text-align` + docx `w:jc` + md 写法）。
3. **「节」工具条的控件要光标落进正文才出现**（`sectionCtx` 为空时只有一句提示，外壳是常驻的）。
   修法：让组件在没有落点时也回一个「首节」的上下文（先定清楚这时候显示什么），
   首节的「关联前节」「从 1 开始」照常置灰。`verify-editor` S 节与 `verify-browser` 第 8 节各有一条断言（红着）。
4. **`--- 文字` 会被静默吃字**：`SECTION_BREAK_RE` 放宽成 `/^-{3,}(?:\s+(.*))?$/` 之后，`--- 备注` 这种行
   被当成带 kwargs 的分节符、其后的文字被丢掉（在加 kwargs 之前它是普通正文）。UI 产不出这种行
   （编辑器只写 `---` 或不写），只有手写 md 才会踩到。要修就把判据收紧成「`---` 尾部只能是一串已知的
   `key=value`」，其余情况**落回正文** —— 与 `parse.ts` 里「围栏没闭合绝不当表格、绝不吞后文」是同一条原则。
5. **相邻两个表格块在 docx 里会被 Word 合并成一张**（模型里中间没有任何段落时）。实现代理实测：
   不隔开时 Word 报「表数 1、行数 12」；样本是靠手写一段过渡文字绕开的。要修就二选一：
   导出时在相邻表之间补一个空段落，或把「连续表块」的模型语义直接改成合并。
6. **md 参差行补出的 `<td>` 是幻影格**：md 里后行比首行多格时，渲染会补空 `<td>`，但模型里没有这些格
   —— 在里面打字不落模型、下次重建即丢。键盘跨格（`stepCell`）现在会跳过幻影格，
   但**在幻影格里打字仍然不落模型**（`findContainer` 找不到那些格）。
7. **`columns` 只能由数据行回推**：手搓模型里 `columns` 大于所有数据行的格数时，md 往返会把它收窄
   （md 语法没有表达「总列数」的地方）。真实路径（手写 md、UI）不受影响；真要修就给围栏加个 `columns=`。
8. **三处表格特性没进 Word 对账**（模型 / md / docx 三层都已支持，只是样本里没有、`assert-docx` 就验不到）：
   ① **表格里的批注锚点** —— 为避开 `assert-docx` 里「批注 1 条」那条既有断言；
   ② **格内多段落**（`{p}`）—— 字节层已验（`verify-docx` 另造一份最小 docx），但 Word COM 那一侧读不到。
   补 ② 的做法是在样本表里加一格 `{p}`，**但那会牵动一批按样本格子文字定位的断言**：
   2026-09-16 试过给 `数控加工中心` 那格加 `{p}`，`verify:p3` 当场红 28 条（
   `caretAtEndOf('数控加工中心')` 落在新多出来的第 0 段、段数/行高/垂直对齐的期望全偏）；
   要补就得先把那些断言的定位方式换成「格子 + 段落下标」，不是加一行样本就完事。
   ③ **跨页重复标题行**（`headerRows`）—— 字节层验了（`verify-docx` 逐行 `w:tblHeader`）、
   预览重复渲染与分页算术都有断言（含「一片绝不重复它自己含有的行」这条不变式），
   但 `check-docx.ps1` 报回来的行对象里**没有 `Row.HeadingFormat`** —— Word 自己认不认这一行是标题行，
   没被读到。补它要动三处：`ps1` 的行对象加一行 `headingFormat = [bool]$row.HeadingFormat`（纯 ASCII，安全）；
   样本表标 `headerRows`（实测样本页数会 4→5 —— 样本表本来就跨页，`> 单位：元` 那行单独落在页底）；
   那 4~5 条按「DOM 行数 = 模型行数」写死的断言改成「= 模型行数 + 重复行数」的口径（T 节三条、V1 一条）。
   **用户 2026-09-16 明确选择本轮不动样本**（怕牵动一批按样本定位的断言、也不想改样本页数），先记在这里。
   当前 `assert-docx.mjs` 已写成「模型里没有 `headerRows` 就不判失败」，运行时另打一行
   「dump 里没有 HeadingFormat —— 这一项没验」。
   ⚠️ 要真验「Word 也把重复行画在续页顶端」，光读 `HeadingFormat` 还不够，得读每行落在第几页
   （ps1 里加 `Row.Range.Information(3)` 之类），成本更高；不过一旦样本标上 `headerRows`，
   `verify:pages` 的「预览页数与 Word 逐套对得上」本身就是这条功能的硬证据 —— 重复行高算错就对不上。
9. **半截 `[[`（未闭合的批注括号）**会把其后的 `| b |` 一起吞进同一格 —— 内容不丢、往返稳定，
    只是格数比旧算法少一个，记录备查。
10. **`formatAmount` 的分组校验偏松**：正则首段是不限长的 `\d+`，`1234,567`、`12345,678` 这类非规范分组
    也被接受（结果仍会规范化，不会出错值）；要收紧就改那一处正则。
11. **下划线只认 `<u>` 标签与行内 `style.textDecorationLine`**，纯 CSS 类画的下划线不会被 `readInlines` 读回。
    当前无害（`render/html.ts` 固定用 `<u>`），但**将来若把下划线的渲染改成 class 就会静默丢格式**。
12. **两处用户 2026-09-16 决定「先不修」的**（记着即可）：
    ① **四页功能区只在宽视口等高** —— 相等是靠 `.panel { min-height: 32px }` 顶出来的，
    1700px 宽时 `[49,49,49,49]`（spread 0），900px 宽时 `[87,49,49,81]`。要真正做到任何宽度都等高，
    得照 Word 那样给功能区一个固定高度（窄宽度时折叠分组），那是另一件事。
    ② **`::editor` 的开关写回不走撤销栈**，而 `undo()` 恢复的是整份快照，于是
    「开修订 → 编辑 → 关修订 → `ctrl+Z`」之后**模型**又变成 `trackChanges=true`，而顶栏复选框仍是未勾选
    —— 界面与模型不同步。修法二选一：把 `editor` 排除在撤销快照之外（倾向这条），
    或撤销后让组件补发一次 `editor-flags`。
13. **整洁度（不影响结论）**：`scripts/verify-browser.mjs` 里硬编码了并排间距 `18px`（数值源自
    `render/css.ts` 的 `gap`）；切模板的断言分散在 `verify-browser` 与 `verify-editor` 两个脚本里；
    `@page wtp-portrait` 这条命名页规则现在没有页会挂它（留着不碍事）。
14. **手搓模型的边界**：`verify:p2` / `verify:p3` / `verify:pages` 里凡是「本机浏览器没跑成」时写的断言，
    事后都要单独复核一遍真跑的状态 —— 这个仓库有过「写了没跑、语法错误藏了很久」的先例。
15. **关掉修订模式后，在「刚被标成插入修订」的文字后面继续打字，新字会并进那个 `w:ins` 区间**
    （读回时按 DOM 上下文继承格式 —— 浏览器就是把字插进了那个 `<span>` 里）。在别的段落里打字不受影响。
    这是「DOM 是手感的真相」付的代价，没有做逐字 diff；要修得做逐字 diff，有回归风险。
    `scripts/verify-editor.mjs` 里有一条断言踩在这条路径上。

---

## 9. 本机环境坑（踩过的，照做能省很多时间）

1. **宿主 Qwen Code 会崩在 TUI 的 React 渲染循环上**。调试日志
   `C:\Users\18082\.qwen\debug\<session-id>.txt` 里有 `[STARTUP] [UNCAUGHT_EXCEPTION] React error #185`，
   栈全在 React 的提交/状态更新路径，即终端 UI 层崩，**与内存无关**。触发面比「大块输出」更宽：
   「后台任务完成 → 自动续跑/刷新界面」这条路径本身也不稳。由此的硬纪律：
   ① 主 session 读取一律**有界**（`git status --short` / `git diff --stat` / 定向小段 / 过滤后 ≤25 行），
   子进程报告**不整段转述**，主 session 自己的回复也必须短；② 想彻底绕开那条路径，让子进程完全脱离
   （见第 10 节）。完整的排查与规避开关（`ui.shellOutputMaxLines`、`tools.truncateToolOutput*`、别按 `Ctrl+O`）
   见 `C:\Users\18082\.qwen\qwen-tui-large-output-crash.md`。**长任务做完后另开新 session**。
2. **命令环境**：Windows 11 + cmd.exe（不是 bash），路径统一用正斜杠。
3. **`npm run verify` 很重**：`verify:p1` 会真的用 Word COM 打开 docx 读回格式（本机已装 Word），
   跑一次全量约十几分钟，别在循环里反复跑。`verify:p2` / `verify:p3` 的无头浏览器**用 Chrome**
   （可用 `WTP_BROWSER` 覆盖），**禁止用 Edge**。`check-docx.ps1` 那套脚本必须保持纯 ASCII。
4. **prettier**：仓库 `format` 脚本用 `npx prettier --experimental-cli`。⚠️ **不要对既有文件跑
   `prettier --write`**：根 `.prettierrc.json` 写的是 `tabWidth: 4`，而源码实际是 **2 空格缩进** ——
   一旦 `--write`，整批文件会被按 4 空格重排（2026-09-13 踩过：9 个代码文件被重排，靠
   `git -c core.autocrlf=false checkout --` 加格式化前的工作区快照才逐字节还原）。
   它也默认会对**你根本没碰过**的既有文件报格式问题，那属既有状态。
   另外 **prettier 会忽略 `.qwen/` 下的文件**：把 HEAD 副本拷进 `.qwen/tmp/` 再 `prettier --check`
   会得到假的「全部合规」，不能拿它判定某个文件是否 prettier-clean。
5. **字体依赖本机安装**：预览与 docx 都写「仿宋 / 华文中宋 / Times New Roman」，都是商业字体、不能内置。
   「方正小标宋」这类公文字体本机大概率没有 —— 涉及新字体前先确认，否则预览会回退、行宽会变。
6. **派发用的辅助脚本（`.ps1` / `.cmd`）必须是纯 ASCII**。实测：无 BOM 的 UTF-8 `.ps1` 里只要有中文
   （注释或 `D:\Work\代码\...` 这样的路径），**Windows PowerShell 5.1 会按 ANSI 读**，脚本被读坏 ——
   表现为 `Join-Path` / `Test-Path` 拿到 `$null`、每轮循环刷屏报错，而**同一轮里它仍在正常轮询**，
   所以从退出码上看不出来。教训：等待循环脚本用**纯 ASCII + 相对路径**（cwd 就是仓库根），
   结论/日志路径写死成 `.qwen/tmp/xxx.md`，输出只留一行标记。
7. **日志是 UTF-8**：用 `read_file` 看，**不要用 PowerShell 的 `Get-Content`**（默认 GBK，会花屏）。
8. **`.qwen` 这种点开头目录 `glob` 工具看不见**，查产物要用 shell 的 `dir /b .qwen\tmp\...`，
   别被「找不到」误导。
9. **子进程启动会打一条「MCP server(s) failed to start」**（那三个 `qwen-mm-plugins-*` 起不来），
   **属正常**，内置工具照常可用，不要去修它。

---

## 10. 派发子进程的工作方式

**调度方式**：每一项工作**先派实现代理（独立的 headless CLI 进程）→ 回来再派独立验收代理**，
验收过了才 `git commit`，然后才进下一步。**为什么不用 subagent**：宿主 TUI 会崩在「大块输出渲进 TUI」上
（见第 9 节第 1 条），而 subagent 的报告走的正是渲染进主 session 这条路径 —— 三次崩溃全部落在报告落地那一刻。
headless 进程没有 TUI，把 stdout 重定向到文件之后，进主 session 的只有 shell 工具那几行摘要。

**派发实现代理时，必须写进任务里**：

1. 同步扩 `scripts/` 下对应的验收脚本（这是本仓库的既有规矩：每一条功能都带断言，见第 6 节）；
2. 收尾必须跑 `npm run type-check` 与相关 `npm run verify:pX`，**全绿才回报**；
3. **不要 commit**（提交由主 session 在验收通过后做）；
4. **结论单独写进 `.qwen/tmp/wX-impl-result.md`（≤ 40 行）**：改了哪些文件、跑了哪些命令与结果、
   「没做到 / 不确定」的部分（不许美化）；细节与完整日志留在 `wX-impl.log`。
   主 session 只读那个结果文件，需要细节再 grep 日志。

**派发验收代理时，必须写进任务里**：

- 仓库内**任何文件都不许改**（发现缺陷只报告不修）；需要临时脚本只许写在 `.qwen/tmp/` 下；
- **长命令输出一律重定向到文件**（`%TEMP%\` 或 `.qwen/tmp/`），只回摘要；
- 结论写进 `.qwen/tmp/wX-review.md`（≤ 40 行：逐条主张给「确认 / 被证伪 / 无法验证」+ 一个总判定），
  证据与命令输出放同一个文件的后面；主 session 只读小窗口；
- 任务是**证伪**实现者的自报，不是复述；
- **不允许自行唤起 office word**。要验证 docx 是否正确，可以交给用户，或者自己解压看 xml。

**命令形态**：

```
qwen -p "Read .qwen/tmp/wXn-task.md as UTF-8 and execute it exactly as written." -y > .qwen/tmp/wXn-impl.log 2>&1
```

- **不批准 shell**：非交互 session 里 `run_shell_command` 会被直接拒掉（日志里只有一句
  `Warning: Tool "run_shell_command" requires user approval but cannot execute in non-interactive mode`）。
  所以任务书里凡含「跑 npm / 起浏览器」一律给 `-y`，主 session 派发后要自己先跑一次 `type-check` 兜底。
- 若这一轮既不用跑命令也不用写脚本，用 `--approval-mode plan` 替换 `-y`，把写操作拦在执行前，
  而且不会挂住（写完「待批准方案」就退出）。
- 可选护栏：`--max-session-turns N`（超限退出码 53）、`--max-wall-time`、`--max-tool-calls`（预算类退出码 55）。
- **想彻底绕开「后台任务完成 → 自动续跑」那条不稳的路径**，用完全脱离的起法：
  把启动命令写成 `.qwen/tmp/wXn-run.cmd`（内里是上面那条 `qwen -p … > … 2>&1`），再用
  `Start-Process -FilePath cmd.exe -ArgumentList '/c','.qwen\tmp\wXn-run.cmd' -WindowStyle Hidden` 起。
  ⚠️ **不要用 `start /b`**：它共享控制台，shell 工具会把整条命令等满超时（默认 120 s），
  而超时又按进程组把子进程一起 `^C` 掉（日志里只剩一句 `^C`）。
  代价：失败 / 挂住没有信号，只能靠轮询结论文件 + 判活发现 ——
  **判活的可靠做法**是查命令行里含任务书文件名的进程是否还在
  （`Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%wXn-task.md%'"`），
  注意命令行的 `-p` 只给一句纯 ASCII 指针时不会自匹配。
- **「日志没长」不等于进程死了**：`-p` 模式下的 stdout 要等最终产出才落盘（一轮日志可能只有几 KB）。
- 这条起法会被 `auto` 批准模式**拦一次**（判为「隐藏窗口 + 脚本间接调用 `qwen -y` = 绕过权限判定」）。
  按规则**不换路径绕、原样重发同一条命令**即走人工批准，放行后正常脱离运行。

**为什么并行度低**：这个仓库的功能横切面很集中 —— 工具栏在组件里，编辑手感与分页渲染在
`WordPaper.vue`（约 1300 行），模型/渲染/导出又在 `src/lib/*` 横穿。任意一组成员功能的文件集都高度重叠，
多代理同写同一批文件最后要手工合并，得不偿失。**默认串行；只有「新增独立纯函数模块」这类真正不重叠的活儿才并行。**

**提交风格**：中文 commit message，多条改动用 ①②③ 分段讲清「为什么」，结尾标注对应的 issue。一次提交装一件事。
