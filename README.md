# OfficeKit

**让 Agent 交付真正的 Office 与 PDF 文件。**

**简体中文** | [English](README.en.md)

创建报告、搭建 Excel 模型、制作演示文稿、处理 PDF。OfficeKit 把创建、导入、
编辑、渲染和验证放进同一条可追踪的工作流，输出的是可以继续打开、修改和检查的文件，
不只是聊天文本或一张预览图。

OfficeKit 面向使用 Agent 构建自动化工作流的开发者和团队。入口是可直接安装的
Skills 和 JavaScript API；Office 由 OpenChestnut C#/.NET WASM 负责，PDF 由懒加载的
MuPDF.js 负责。

> 仓库包名暂为 `open-office-artifact-tool`，`OfficeKit` 是产品名。当前版本为 `0.3.0` release candidate，尚未正式发布到 npm。

## 适合这些工作

- **报告和文档**：把提纲、资料和数据整理成带样式、表格、图片、字段和评论的 DOCX，导出后重新导入并做页面级检查。
- **财务和运营模型**：把 CSV、假设和业务数据变成带公式、验证、条件格式、图表和计算结果的 XLSX。
- **汇报和方案演示**：沿用已有模板或从 Compose 开始制作 PPTX，保持图片裁剪、图表、备注和版式，再渲染成页面检查。
- **PDF 处理**：读取任意 PDF 的文字、表格、图片、链接和表单；需要修复、OCR、严格清理、签名或合规检查时，再按任务启用对应 provider。
- **模板驱动的批量产出**：从一份本地 DOCX、XLSX 或 PPTX 参考文件创建可复用模板，生成新文件，不改动原始模板。

## 先装起来

### 只装 Skills

不需要 clone 仓库，也不需要安装 .NET、Python 或 Office：

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill documents \
  --skill Spreadsheets \
  --skill Presentations \
  --skill pdf \
  --skill template-creator \
  --yes
```

需要所有原生 Skill 和开源模板时：

```sh
npx skills add w31r4/open-office-artifact-tool --skill '*' --yes
```

Skill 会安装到 Agent 能发现的位置，不需要手工复制 `skills/` 目录。需要用户级安装时加
`--global`，需要固定宿主时加 `--agent`。

### 从空目录启动一个完整项目

```sh
mkdir officekit-agent && cd officekit-agent
npm init -y
npx skills add w31r4/open-office-artifact-tool --skill documents --skill Spreadsheets --skill Presentations --skill pdf --yes
npm install github:w31r4/open-office-artifact-tool
```

正式发布后，把最后一行换成 `npm install open-office-artifact-tool`。

消费者只需要 Node.js；npm 包包含 OpenChestnut WASM，不需要本机 .NET。建议使用 Node.js 22
或更新版本。MuPDF 只在第一次执行 PDF 操作时加载，安装包和普通 Office 操作不会下载
额外运行时。

### 只试用一个 Skill

```sh
npx skills use w31r4/open-office-artifact-tool --skill pdf
```

## 从 JavaScript 调用

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

可运行示例：

- [创建 DOCX 报告](examples/create-docx-report.mjs)
- [创建 XLSX 仪表盘](examples/create-xlsx-dashboard.mjs)
- [使用 Compose 创建 PPTX](examples/create-pptx-compose.mjs)
- [解析与渲染 PDF](examples/parse-render-pdf.mjs)

## 四类文件，一套交付标准

| 格式 | 默认引擎 | 主要用途 |
| --- | --- | --- |
| DOCX | OpenChestnut C# WASM | 结构化文档、样式、分节、表格、图片、字段、评论和有界内容控件。 |
| XLSX | OpenChestnut C# WASM | 单元格、公式、样式、布局、表格、图片、验证、条件格式、评论、图表、sparklines、What-If 数据表和有界 PivotTable。 |
| PPTX | OpenChestnut C# WASM | 形状、富文本、可逆图片裁剪、表格、连接线、图表、备注、评论、Master/Layout 保真和有界源绑定编辑。 |
| PDF | PdfArtifact + MuPDF.js | Tagged PDF 创建；读取、检查、渲染、表单、链接、批注、页面编辑、rewrite 脱敏和有界签名。 |

Office 的普通导入和导出只有 OpenChestnut 一条路径。无法安全建模的内容会保留原始包；
不支持的修改会明确失败，避免生成看似成功但内容已经变化的文件。

## PDF 专项能力按需启用

`mupdf@1.28.0` 是 PDF 的必需 npm 依赖，随正常 `npm install` 解析并在运行时懒加载。
qpdf、Python、OCR、veraPDF/JRE 等专项工具不进入 npm tarball，也不会通过 lifecycle hook
或全局包管理器自动安装。

让 Agent 先根据任务和文件检查结果选择 provider：

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

默认策略是 `disabled`。只有项目明确授权 `managed`，并满足平台、哈希、体积、许可证和
语言包限制时，才能 `ensure`；已有部署也可以明确选择 `system-only`。没有隐式 fallback。
完整策略见 [PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md)。

## Skills 和模板

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)
- [Template Creator](skills/template-creator/skills/template-creator/SKILL.md)
- [Office Template Library](skills/default-template-library/README.md) — 20 套 MIT 授权模板，仅仓库分发，不进入 npm runtime tarball。

模板和 Office 引擎是同一套交付链的一部分：Agent 选定模板后生成新的输出文件，继续用
同一套 API 检查、编辑、渲染和验证，参考文件本身不会被覆盖。

## 交付前检查

```text
intent → inspect → resolve → edit/create → export → re-import → render → verify
```

源快照丢失、拓扑不可信、签名会失效、PDF 操作可能保留旧 revision，或缺少外部凭据时，
操作会 fail closed。完整能力边界见 [coverage](docs/coverage.md)。

## 开发与验证

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

API 参考：[docs/api.md](docs/api.md)；参考 Skill 兼容性：[docs/reference-skills.md](docs/reference-skills.md)。

## 许可证

[GNU AGPL v3 或更高版本](LICENSE)。网络部署、修改和分发必须遵守 AGPL 的对应义务。
第三方运行时、MuPDF 和专项 provider 的许可证与来源见 `THIRD_PARTY_NOTICES.md` 及相关 runtime notices。
