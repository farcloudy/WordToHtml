# WordToHtml

把「类 md 字符串」编译成 **A4 分页预览** 与 **Word 可编辑的 docx**，并按 Word 的样式名把 docx 读回来。

面向中文正式公文（情况说明、报告、立场文件）这一「特定模式」：字体、字号、行距、段距、
标题编号规则全部按固定规格表来，规格表是数据，可以整表覆盖，没有一处硬编码。

## 当前进度

| 阶段 | 内容 | 状态 |
| --- | --- | --- |
| P1 | 规格表 → md 解析 → docx 导出（页码／分节／修订／批注） | ✅ 已验收（Word 逐项对账通过） |
| P2 | A4 分页预览（自研分页器 + 页码 + 分节重编号） | ✅ 已验收（页数与 Word 一致） |
| P3 | 编辑层（所见即所得、选中加粗改色、修订模式、批注 UI） | ⬜ 未开始 |
| P4 | docx 反向导入（按样式名匹配） | ⬜ 未开始 |

## 快速开始

```bash
npm install
npm run dev        # 打开 demo：左侧写源码，右侧即时出 A4 版面，右上角导出 docx
npm run verify     # 跑完整验收（类型检查 + Word 对账 + 分页单测 + 浏览器实测 + 页数比对）
```

## 组件用法

```vue
<script setup lang="ts">
import { ref } from 'vue'
import WordPaper from './components/WordPaper.vue'
import { mm } from './lib/spec'
import type { DeepPartial, Spec } from './lib/spec'

const source = ref('# 关于××的情况说明\n\n@ 苏州市公安局：\n\n……')
const paper = ref<InstanceType<typeof WordPaper> | null>(null)

// 任何一项都能覆盖，页边距只是举例
const spec = ref<DeepPartial<Spec>>({
  page: { margin: { top: mm(30), right: mm(25), bottom: mm(30), left: mm(25) } },
})
</script>

<template>
  <WordPaper ref="paper" :source="source" :spec="spec" author="张三" @paginated="n => console.log(n)" />
  <button @click="paper?.downloadDocx('情况说明.docx')">导出 docx</button>
</template>
```

`defineExpose` 暴露：

| 方法 | 说明 |
| --- | --- |
| `repaginate()` | 强制重新量测与分页（一般不需要手动调） |
| `exportDocx(): Promise<Blob>` | 取 docx 二进制 |
| `downloadDocx(name?)` | 直接触发浏览器下载 |
| `getModel()` / `getSpec()` | 取当前模型／生效规格 |
| `getMeasurements()` | 取最近一次分页用到的量测值（行数、行高、段距），排错用 |
| `pageCount()` | 当前页数 |

props：`source`（类 md 源码）、`model`（直接给模型，优先于 source）、`spec`（规格覆盖）、`author`（修订与批注作者名）。
事件：`paginated`（分页完成后给出页数）。

文档含批注时，组件右侧自动出现审阅侧栏（锚定文字 + 作者 + 时间 + 内容），点条目会在正文里
高亮对应锚点并滚动过去；没有批注时侧栏不占位。页边距预设见 `MARGIN_PRESETS`，取 `preset.margin`
传给 `spec.page.margin` 即可整组切换。

## 源码语法

一个非空行就是一段（**行与行不会合并成同一段** —— 公文写作习惯是一行一段，用空行分隔反而要多敲回车）。
空行只起分隔作用，不产出内容。

| 行首标记 | 含义 |
| --- | --- |
| `# 标题` | 文本标题（华文中宋 18pt 加粗居中） |
| `## 标题` | 一级标题，自动编号 `一、` |
| `### 标题` | 二级标题，自动编号 `（一）` |
| `#### 标题` | 三级标题，自动编号 `1、` |
| `@ 抬头` | 抬头（正文但取消首行缩进） |
| `>> 落款` | 落款（正文但右对齐），连续多行成组 |
| `- 条目` | 列表段落（五号 10.5pt，无段前段后） |
| `! 条目` | 列表标题（五号加粗居中） |
| `---`（独占一行） | 分节符：新起一页，页码从 1 重排 |

行内标记（可互相嵌套）：

| 写法 | 含义 |
| --- | --- |
| `**文字**` | 加粗 |
| `{红｜文字}` | 改色，支持中文色名（红黑蓝绿紫橙灰黄白）或 `{#RRGGBB｜文字}` |
| `{+文字}` | 插入修订（导出为 `w:ins`） |
| `{-文字}` | 删除修订（导出为 `w:del`） |
| `[[文字｜批注内容]]` | 批注，锚定在「文字」上 |
| `\` | 转义上述标记字符 |

> 上表里的全角竖线是为了在文档里显示清楚，实际语法用的是**半角 `|`**。

## 样式规格表

规格表在 [`src/lib/spec.ts`](src/lib/spec.ts)，是预览 CSS 与 docx 样式的**唯一来源**，
`resolveSpec(override)` 做逐层合并。默认值：

| 类型 | 字体（中文／西文） | 字号 | 加粗 | 对齐 | 首行缩进 | 行距 | 段前／段后 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 文本标题 | 华文中宋／Times New Roman | 18pt | ✓ | 居中 | — | 26pt 固定值 | 1.5 行／1.5 行 |
| 一~三级标题 | 仿宋／Times New Roman | 14pt | ✓ | 两端 | 2 字符 | 25pt 固定值 | 0.5 行／0.5 行 |
| 正文 | 仿宋／Times New Roman | 14pt | — | 两端 | 2 字符 | 25pt 固定值 | 0.5 行／0.5 行 |
| 抬头 | 仿宋／Times New Roman | 14pt | — | 左 | — | 25pt 固定值 | 0.5 行／0.5 行 |
| 落款 | 仿宋／Times New Roman | 14pt | — | 右 | — | 25pt 固定值 | 0.5 行／0.5 行 |
| 列表标题 | 仿宋／Times New Roman | 10.5pt | ✓ | 居中 | — | 12pt 最小值 | 0／0 |
| 列表段落 | 仿宋／Times New Roman | 10.5pt | — | 两端 | 2 字符 | 12pt 最小值 | 0／0 |

- 页面默认 A4，四边 25mm。`MARGIN_PRESETS` 提供「四边 25mm」与「公文标准（上37 下35 左28 右26）」
  两组预设，四边各自取值；`page.margin` 也可以直接覆盖。页码段落用内置「页脚」样式（居中 9pt）。
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

## 架构

```
src/lib/
  spec.ts            页面与样式的唯一真相源 + 页边距预设 + 单位换算
  types.ts           文档模型（块 / 行内 / 修订 / 批注）+ 批注锚定文字提取
  numbering.ts       中文序数编号与层级重置
  md/parse.ts        类 md → 模型        md/serialize.ts  模型 → 类 md
  docx/export.ts     模型 → docx（段落样式、页码域、w:ins/w:del、批注）
  render/css.ts      规格表 → 预览 CSS
  render/html.ts     行内标记 → HTML（预览与量测共用同一函数，保证量到即看到）
  render/measure.ts  DOM 实测：行数、行高、每行起始字符偏移
  render/paginate.ts 纯函数分页：装箱 + 跨页按行切开 + 分节重编号
src/components/WordPaper.vue   预览组件 + 批注审阅侧栏
scripts/                       验收脚本（见下）
```

分页不是 CSS 断页，而是「实测行盒 → 装箱 → 在行边界切开」算出来的：行距是固定值，
行高因此是确定量，分页可以算出来而不是猜。孤行控制默认开启（页尾与页首各至少 2 行，
因此少于 4 行的段落不会被拆开）。

## 验收方式

| 命令 | 验的是什么 |
| --- | --- |
| `npm run verify:p1` | 生成 docx → **Word COM 打开** → 逐项读回 Word 实际生效的字体／字号／行距规则／段距／首行缩进字符数／样式归属（含 Word 是否把它认成内置样式）／页码域／修订／批注，与规格表对账 |
| `npm run verify:p2` | 分页器 39 项单测 + 真实浏览器（系统 Edge）实测：每页不得溢出、样式与规格表一致、页首与续排的间距豁免、**量测值与渲染值逐块对账**、批注侧栏与正文锚点同源（含点击高亮） |
| `npm run verify:pages` | 同一份源码分别交给预览与 Word，比对页数与分节重编号是否一致 |
| `npm run verify` | 以上全部 + 类型检查 |

`verify:p1` 会在机器上启动 Word（只读打开、读完即退出）。若本机没装 Word，这一步会失败，
其余步骤不受影响。

## 已知限制

- **字体依赖本机安装**。预览用 `仿宋／华文中宋／Times New Roman`，docx 里写的也是这几个名字。
  这三个都是商业字体，不能内置；换一台没装这些字体的机器，预览会回退到 serif，行宽会变。
- **「段前／段后 N 行」用磅值写死**。`docx` 库没有暴露 Word 的 `w:beforeLines`（按行计的动态段距），
  这里按本段行距换算成固定磅值。视觉等价，但不是随行距联动的动态值。
- **自动编号写成了正文文字**，不是 Word 原生多级列表。好处是预览所见即所得、重新导入能准确还原；
  代价是在 Word 里手动增删标题不会自动重编号（在本组件内会自动重排）。
- **分节符会多出一个空段落**。`docx` 库用一个只带 `w:sectPr` 的空段落承载分节符，
  Word 把它读成含分页符的一段，效果是上一节末尾多一个空行。
- **批注的「已解决」状态需要线程**。`docx` 只在存在 `parentId` 回复时才会写
  `commentsExtended.xml`，因此单条批注标记 `resolved` 不会生效（库的现状，非本项目取舍）。
- 预览里的批注高亮与修订着色是**近似**，不追求与 Word 像素级一致。
- 预览分页与 Word 分页在同样规则下目前完全一致（见 `verify:pages`），但两者毕竟是不同的排版引擎，
  换字体、混排西文比例很高、或出现超大表格时仍可能出现逐行差异。

## 待做

- **P3 编辑层**：`WordPaper` 目前是只读的。已定的两点：
  - 编辑形态：**直接在 A4 版面上编辑**（contenteditable 落在分页后的页面上，而不是另开一个编辑区）；
  - 关键约束：**仅当分页结果真的变化时才重排并恢复插入符** —— 正常输入不得触发重排，否则每敲一个字就丢光标。

  其余要补的是工具栏（加粗／改色）、修订模式开关，以及批注的新增与回复（侧栏已在，缺写入口）。
- **P4 docx 反向导入**：解压 → **按样式名反查 `BlockKind`**（内置名与自定义名同一张表，见上）→
  读 `w:ins`/`w:del` 与批注锚点 → 还原模型（`stripAutoNumber` 已备好，用于剥掉段首已有的编号避免重复编号）。
  注意正文：导出时它不写 `w:pStyle`，导入时要把它当 `body`，不能当异常跳过。
