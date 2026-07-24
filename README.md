# OfficeKit

## 让 Agent 交得出 Word、Excel、PPT 和 PDF

**简体中文** | [English](README.en.md)

写一段内容很容易。把它做成明天要开会、发给客户、留作正式记录，下一次还能继续修改的文件，要求高得多。

OfficeKit 给 Agent 一套处理 Word、Excel、PowerPoint 和 PDF 的 Skills 与 JavaScript API：创建新文件，读取已有文件，做受支持的局部修改，再把结果重新打开、渲染并检查。无法可靠保留的复杂内容不会被悄悄改坏；有风险的修改会明确停下来说明原因。

## 直接说出你要交付的东西

把下面这类任务交给装有 OfficeKit 的 Agent：

> “把这几份 CSV 做成下周经营会用的 Excel 模型；关键指标保留公式，图表可以直接放进汇报。”

> “沿用公司的 PPT 模板做一套 QBR。替换数据和图片，检查每一页有没有溢出或错位。”

> “更新这份 Word 里的日期、负责人和条款，页眉、目录、引用和现有批注不能乱。”

> “把这批扫描 PDF 变成可搜索文件，标出敏感信息、做脱敏，并给我一份可核查的结果。”

OfficeKit 既能从零开始生成，也能在原文件上工作。它适合报告、财务模型、客户方案、培训材料、合同草稿、批量模板和 PDF 处理任务。

## 两步装进项目

在 Agent 要工作的 Node.js 项目里执行：

```sh
npm install github:w31r4/open-office-artifact-tool
npx skills add w31r4/open-office-artifact-tool --skill '*' --yes
```

第一行安装运行库，第二行安装全部 Skills 和开源模板。首次接入无需 clone 仓库，也不需要安装 Office、.NET 或 Python。当前正式 npm 包尚未发布，因此先使用 GitHub 安装源；发布后可换成 `npm install open-office-artifact-tool`。

只需要部分能力时：

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill documents \
  --skill spreadsheets \
  --skill presentations \
  --skill pdf \
  --yes
```

推荐 Node.js 22 或更新版本。Office 运行时已随包提供；PDF 的 MuPDF.js 只会在第一次 PDF 操作时加载。

## Skills 让 Agent 知道怎么做

Skills 不是一份功能清单。它们告诉 Agent 先看什么、改完如何检查、什么情况必须停下：

| Skill | 适合的任务 |
| --- | --- |
| [Documents](skills/documents/skills/documents/SKILL.md) | Word 报告、函件、合同草稿、带表格和图片的正式文档。 |
| [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md) | Excel 模型、数据整理、公式、图表、验证和可视化。 |
| [Presentations](skills/presentations/skills/presentations/SKILL.md) | PowerPoint 汇报、模板套用、图表、图片、备注和版式检查。 |
| [PDF](skills/pdf/skills/pdf/SKILL.md) | PDF 读取、创建、表单、批注、页面处理、渲染和专项处理。 |
| [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) | 从自己的 DOCX、XLSX、PPTX 参考文件制作可复用模板。 |

Skills 和应用代码使用同一个包。Agent 可以按 Skill 完成任务，应用也可以直接调用 API，把文件能力接进自己的产品或自动化任务。

## 从代码调用

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");
sheet.getRange("A1:B2").values = [
  ["Metric", "Value"],
  ["Revenue", 42.5],
];

const file = await SpreadsheetFile.exportXlsx(workbook, { recalculate: true });
await file.save("summary.xlsx");
```

可直接运行的例子：

- [创建 DOCX 报告](examples/create-docx-report.mjs)
- [创建 XLSX 仪表盘](examples/create-xlsx-dashboard.mjs)
- [使用 Compose 创建 PPTX](examples/create-pptx-compose.mjs)
- [解析与渲染 PDF](examples/parse-render-pdf.mjs)

## 交付前，文件会被再看一遍

OfficeKit 的常用路径很简单：

```text
读取原件 → 创建或修改 → 导出 → 重新打开 → 渲染页面 → 检查结果
```

- DOCX、XLSX 和 PPTX 统一走 OpenChestnut C#/.NET WASM。没有第二套 JS Office writer 在背后兜底，也不需要本机 .NET。
- 对无法安全建模的 Office 内容，系统会尽量保留原文件中相应部分；请求会破坏它时，修改会失败，而不是输出一个看起来正常、实际已变形的文件。
- PDF 由 MuPDF.js 处理默认读写、检查和渲染。修复、OCR、严格清理、签名、PDF/A 或 PDF/UA 等专项能力需要明确选择对应 provider。

| 文件 | 常见交付内容 |
| --- | --- |
| DOCX | 样式、段落、分节、页眉页脚、表格、图片、字段、评论，以及对现有文档的局部编辑。 |
| XLSX | 单元格、公式、样式、合并、尺寸、冻结、表格、图片、数据验证、条件格式、图表、sparklines 和有界 PivotTable。 |
| PPTX | 形状、富文本、图片与可逆裁剪、表格、连接线、图表、备注、评论和 Master/Layout 保真。 |
| PDF | 创建、提取文本/表格/图片/链接、表单和批注、页面操作、rewrite 脱敏，以及有界签名。 |

完整支持边界见 [coverage](docs/coverage.md)。

## PDF 的重型能力，按项目授权启用

MuPDF 是正常依赖。qpdf、Python、OCR、veraPDF/JRE 等不会塞进 npm 包，也不会由安装脚本或全局包管理器悄悄下载。

Agent 先根据任务和文件检查结果选择 provider；项目默认不允许下载。只有在 `.open-office-artifact-tool/pdf-providers.json` 中明确把策略设为 `managed`，并且平台、哈希、体积、许可证和 OCR 语言都符合约束时，才可以安装受管能力包。已有运行时也可以明确选择 `system-only`。

```js
import { PdfFile } from "open-office-artifact-tool";
import { PdfProviders } from "open-office-artifact-tool/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");
const resolution = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  inspection,
});

console.log(resolution.status); // ready | installable | blocked
```

[PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md) 说明了策略、安装和各项能力的限制。

## 模板：从已有好文件开始

[Office Template Library](skills/default-template-library/README.md) 提供 20 套 MIT 授权的 Office 模板。模板文件留在仓库中，不进入 npm runtime 包；Agent 基于选定模板生成新的输出，再用同一套 API 检查、编辑和渲染，参考文件不会被当作输出覆盖。

也可以用 [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) 把团队现有的 DOCX、XLSX 或 PPTX 参考文件变成自己的模板。

## 给使用者和贡献者

- [API 参考](docs/api.md)
- [参考 Skill 兼容性](docs/reference-skills.md)
- [全部能力边界](docs/coverage.md)

开发时运行：

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

> `OfficeKit` 是产品名；当前包名仍为 `open-office-artifact-tool`。版本 `0.3.0` 为 release candidate，尚未正式发布到 npm。

## 许可证

[GNU AGPL v3 或更高版本](LICENSE)。网络部署、修改和分发必须遵守 AGPL 的对应义务。第三方运行时、MuPDF 和专项 provider 的许可证与来源见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) 及相关 runtime notices。
