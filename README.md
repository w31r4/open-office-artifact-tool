# open-office-artifact-tool

**简体中文** | [English](README.en.md)

面向 Agent 的 Office 与 PDF 创建、读取、编辑、检查、渲染和验证工具箱。

`open-office-artifact-tool` 提供统一的 JavaScript 对象模型。DOCX、XLSX 和 PPTX 由仓库内的 **OpenChestnut**（C# + Open XML SDK + .NET WebAssembly）读写；PDF 使用独立语义模型和显式 Provider 路由。

> **当前状态：** `0.2.0` 发布候选。源码、可复现 WASM 和 npm tarball 已具备验证流程，但尚未执行正式 `npm publish`。

## 快速开始

正式发布前，请从源码运行：

```sh
git clone https://github.com/w31r4/open-office-artifact-tool.git
cd open-office-artifact-tool
npm install
node examples/create-xlsx-dashboard.mjs
```

本地发布门禁已在 Node.js 26.5.0 通过，托管 CI 使用 Node.js 22；这是已验证环境，并非已固化的最低版本。普通消费者直接加载仓库或 npm 包内的 WASM，不需要本机 .NET；重建 OpenChestnut，或构建/测试可选的 OfficeBridge，需要 .NET SDK 8。

也可以直接使用公共 API：

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");

sheet.getRange("A1:B2").values = [
  ["Metric", "Value"],
  ["Revenue", 42.5],
];

const xlsx = await SpreadsheetFile.exportXlsx(workbook, { recalculate: true });
const reopened = await SpreadsheetFile.importXlsx(xlsx);
console.log(reopened.inspect({ kind: "worksheet,table,chart" }).ndjson);
```

更多可运行示例：

- [创建 DOCX 报告](examples/create-docx-report.mjs)
- [创建 XLSX 仪表盘](examples/create-xlsx-dashboard.mjs)
- [使用 Compose 创建 PPTX](examples/create-pptx-compose.mjs)
- [解析与渲染 PDF](examples/parse-render-pdf.mjs)

## 为什么需要它

- **为 Agent 设计**：文件模型自带 `inspect`、`resolve`、`verify`、render 和 visual QA 原语。
- **保真优先**：无法安全建模的 Office 内容绑定原始包并原样保留；不支持的修改明确失败。
- **内置原生 Skills**：仓库随包提供 Documents、Spreadsheets、Presentations 和 PDF 四个 Codex 插件包；需要宿主会话或外部 Provider 的工作流会明确说明前置条件。

## 支持范围

| 格式 | 文件管线 | 当前核心能力 |
| --- | --- | --- |
| XLSX | OpenChestnut C# WASM | 单元格与公式、样式与布局、表格、图片、基础验证/条件格式、评论、图表和 sparklines。 |
| DOCX | OpenChestnut C# WASM | 结构化文本与样式、分节、页眉页脚、列表、表格、链接、字段、图片和经典评论。 |
| PPTX | OpenChestnut C# WASM | 形状与富文本、图片及可逆裁剪、表格、连接线、图表、直接背景和纯文本演讲者备注；Master/Layout 仅保真、不可编辑。 |
| PDF | 独立模型与 Provider 路由 | 内置 Tagged PDF 创建、结构与阅读顺序、表格/图片/链接和模型 QA；外部 Provider 承担表单、批注、有界原文件编辑、合并重排、水印、脱敏和原生渲染。 |

完整且持续更新的边界见 [能力矩阵](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/coverage.md)。

## 工作方式

```text
Agent / Codex Skill
├─ Office → JavaScript model → OpenChestnut C# WASM → DOCX / XLSX / PPTX
├─ PDF    → PdfArtifact 或显式 Provider → PDF
└─ QA     → inspect / resolve → render → verify / visual QA
```

OpenChestnut 是普通 Office 导入/导出的唯一 parser/writer。显式 OOXML inspect/patch 只供高级用户手动调用，不会成为自动 fallback。

## 原生 Skills

仓库包含四个插件包、五个 Skill：

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) — 依赖宿主提供实时 Excel 会话
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)

四个 `skills/<name>` 目录都是完整的 Codex 插件根目录；发布前需先创建一个引用这些目录的本地 marketplace，再从该 marketplace 安装插件。Office Skill 的普通文件工作流统一调用 OpenChestnut。PDF npm 表面包含 JavaScript 模型和薄路由脚本，不捆绑外部 Provider 本身；缺少必需依赖时会明确失败。安装与职责见 [PDF Provider Matrix](skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md)。

## 必须知道的边界

- 要保留导入 Office 文件中的未建模对象，必须继续使用 import 返回的模型，并保持这些对象的结构不变；丢失源快照或修改不支持的拓扑时，导出失败。
- 任意已有 PDF 不能像 Word 一样可靠地自动重排全文；原文件编辑必须落在明确、可验证的有界操作中。
- PDF 签名、时间戳与 LTV 依赖外部 pyHanko 工作流，PDF/A 与 PDF/UA 机器验证依赖外部 veraPDF；它们不是随包提供的完整适配器。
- PyMuPDF/MuPDF 不随 MIT npm 包分发。使用者必须单独安装，并明确接受 GNU AGPL 或商业许可。
- LibreOffice、Poppler、Playwright 和原生 Office Bridge 是渲染/验证工具，不是隐藏的 Office codec fallback。

## 开发与验证

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

继续阅读：[API](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/api.md) · [运行时架构](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/reference-runtime-architecture.md) · [Skill 兼容性](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/reference-skills.md) · [Agent PromptBench](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/agent-evals.md) · [发布门禁](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/release.md)

以上文档链接跟随当前开发主线；正式发布时将固定到对应版本标签。

## 许可证

[MIT](LICENSE)。第三方运行时许可与来源见 `THIRD_PARTY_NOTICES.md` 和 OpenChestnut runtime notices。
