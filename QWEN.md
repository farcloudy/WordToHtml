# WordToHtml 开发文档（二次开发参考）

> `README.md` 是给**使用者**的：怎么跑起来、源码怎么写、快捷键、已知限制。
> 这份文件是给**改这个仓库的人**的：约束、已定稿的设计、文件地图、组件接口、验收脚本、还开着的问题、本机环境坑。
>
> **维护约定**：做完一件事就把「过程与结论」从这份文件里删掉 —— 需要考古时 `git log` 与提交信息里都有。
> 这里只留**仍然有效**的东西：约束、接口事实、还没做完的、还没修的。

---

## 1. 踩坑提示

1. **规格表是唯一真相源**。`src/lib/spec.ts` 是预览 CSS 与 docx 样式的共同来源；改样式只改那张表，预览 CSS（`render/css.ts`）、docx 样式（`docx/export.ts`）、量测段距三处都从它派生。
2. **验收脚本的期望值从 `resolveSpec()` 推导，不硬编码数值**。所以改 spec 之后如果断言没跟着变，说明你写错了地方；反过来，样式类改动通常**不需要**改断言。
3. **DOM 是手感的真相，模型是导出的真相**。结构性操作（回车分段、退格合并、改格式、加批注、插入表格…）直接改模型再重排渲染；打字、组字、选区、剪贴板留给浏览器。
4. **docx 的内置样式归属按 `w:name` 逐字匹配 Word 本地化名**（含「标题 1」中间那个空格），不是按 `w:styleId`。`w:styleId` 一律保留 `WT-` 前缀 —— 撞上 `docx` 库的内置表会让它注入重复定义。
5. **命令环境**：Windows 11 + cmd.exe（不是 bash），路径统一用正斜杠。
6. 未经请示同意不得主动使用`npm run verify`跑验证
7. **字体依赖本机安装**：预览与 docx 都写「仿宋 / 华文中宋 / Times New Roman」，都是商业字体、不能内置。
   「方正小标宋」这类公文字体本机大概率没有 —— 涉及新字体前先确认，否则预览会回退、行宽会变。
8. **派发用的辅助脚本（`.ps1` / `.cmd`）必须是纯 ASCII  + 相对路径**（cwd 就是仓库根）
9. **日志是 UTF-8**：用 `read_file` 看，**不要用 PowerShell 的 `Get-Content`**（默认 GBK，会花屏）。
10. **`.qwen` 这种点开头目录 `glob` 工具看不见**，查产物要用 shell 的 `dir /b .qwen\tmp\...`，
   别被「找不到」误导。

---

## 2. 文件地图（接入点）

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

## 3. 组件接口

### 3.1 `WtpEditor`（对外组件）

对外就一个组件：`WtpEditor`（`src/components/WtpEditor.vue`，库里以 named export 导出）。
**顶栏 + 功能区 + 纸张**整副外壳都在它里面 —— 批注侧栏、查找替换面板、插入表格面板、导航窗格、提示条
也归它管。「类 md 源码」pane 与「所见即所得 / 源码」模式切换**不在**组件里（那是 demo 的事，
见下面「插槽」）。

props：

| prop | 类型 | 说明 |
| --- | --- | --- |
| `content` | `string` | 类 md 内容。**它是初始内容**：编辑过程中组件不回写（免得一个字回调一次）。留空 = 空字符串 |
| `model` | `DocModel` | 直接给一份现成的模型。与 `content` **二选一、`model` 优先**（优先级只在 `WordPaper` 那一处 `props.model ?? parseMd(...)` 判）；md 语法表达不了的**修订作者/时间戳**与**批注的作者、回复线程**只有走它才保得住（md → 模型那一步会按 `author` prop 与当前时间重造一份） |
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
- 默认插槽：纸张左侧、导航窗格之后。最小示例：

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

- **`content` 是初始内容**，编辑过程中不回写；要拿最新 md 调组件暴露的 `toMd()`（或 `getModel()`），`save_md` 交出去的就是同一个字符串；
- **`ctrl+S` 归组件管**：按它就等于点「保存」，并且拦住浏览器默认的「保存网页」；
- 组件自带整屏高度（`height: 100vh`）与内部滚动；要嵌在别处就在外面覆盖 `height`。demo（`src/App.vue`）本身就是一份接法示例，另外用 URL 开关演示这几个入口（验收脚本也走它们）：
`?shortcuts=bold:ctrl+shift+b,formatAmount:ctrl+alt+4` 覆盖快捷键表、`?template=govDoc` 直接把模板当 prop
递进去、`?empty=1` 演示 `content` 留空、`?model=empty|rich|rich-md` 演示 `model` 通路
（`empty` = 零块模型；`rich` = 带修订作者/时间戳与批注回复线程的现造模型；`rich-md` = 同一份内容改走
`content` 通路，当 `rich` 的对照）。demo 把组件回传的值（`update:*`、`save_md` 收到的 md）与
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

### 3.2 内部组件 `WordPaper`（高级用法）

详见`内部组件开发说明.md`

---

## 4. 待做

### docx 反向导入

解压 → **按样式名反查 `BlockKind`**（内置名与自定义名同一张表，见第 5 节）→ 读 `w:ins` / `w:del`
与批注锚点 → 还原模型（`stripAutoNumber` 已备好，用于剥掉段首已有的编号避免重复编号）。
**注意正文**：导出时它不写 `w:pStyle`，导入时要把它当 `body`，不能当异常跳过。
