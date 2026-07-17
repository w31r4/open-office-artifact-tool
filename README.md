# open-office-artifact-tool

面向 Agent 的 Office 与 PDF 创建、读取、编辑和验证工具箱。

> A clean-room, agent-facing toolkit for creating, editing, inspecting, rendering, and verifying Office and PDF artifacts.

`open-office-artifact-tool` 用统一的 JavaScript 对象模型提供 Agent 友好的操作原语。DOCX、XLSX 和 PPTX 由仓库内的 **OpenChestnut**（C# + Open XML SDK + .NET WebAssembly）负责真实文件读写；PDF 使用独立的语义模型和显式 Provider 路由。项目不包含第二套 JavaScript Office codec，也不会在失败时偷偷降级为有损输出。

> 当前状态：`0.2.0` 发布候选。源码、可复现 WASM 和 npm tarball 已具备完整验证流程，但尚未执行正式 `npm publish`。

## 为什么需要它

- **为 Agent 设计**：提供 `inspect`、`resolve`、`verify`、render 和 visual QA，而不只是字节级读写。
- **单一 Office 路径**：DOCX、XLSX、PPTX 始终经过 OpenChestnut；安装后的使用者不需要本机 `dotnet`。
- **保真优先**：无法安全建模的 Office 内容会绑定原始包并原样保留；不支持的编辑明确失败。
- **PDF 按能力选路**：创建、提取、表单、原地编辑、脱敏、签名和合规验证分别交给合适的成熟工具。
- **Skill 可直接使用**：仓库随包提供 Documents、Spreadsheets、Presentations 和 PDF 四个原生 Codex 插件包。

## 支持范围

| 格式 | 文件管线 | 当前核心能力 |
| --- | --- | --- |
| XLSX | OpenChestnut C# WASM | 单元格、公式、样式、表格、图片、冻结窗格、基础验证与条件格式、评论、bar/line/pie 图表和标准 sparklines。 |
| DOCX | OpenChestnut C# WASM | 段落与 Run、样式、分节、页眉页脚、列表、固定几何表格、链接、简单字段、图片和经典评论。 |
| PPTX | OpenChestnut C# WASM | 形状、富文本、图片及可逆裁剪、表格、连接线、bar/line/pie 图表、直接幻灯片背景、纯文本演讲者备注，以及 Master/Layout 保真。 |
| PDF | 独立模型与 Provider 路由 | Tagged PDF 创建、提取与阅读顺序、表格/图片/链接、表单与批注、有界原文件编辑、合并重排、水印、真实脱敏、渲染和残留检查。 |

完整且持续更新的边界见 [能力矩阵](docs/coverage.md)。

## 架构

```text
Agent / Codex Skill
├─ Office → JavaScript model → OpenChestnut C# WASM → DOCX / XLSX / PPTX
├─ PDF    → PdfArtifact 或显式 Provider → PDF
└─ QA     → inspect / resolve → render → verify / visual QA
```

JavaScript 负责公共对象模型、计算、Compose/JSX、检查与渲染编排；OpenChestnut 是唯一的 Office parser/writer。PDF 不进入 Office protobuf/WASM 管线，也不会伪装成可任意重排的 Word 文档。

## 快速开始

正式发布前，请从源码安装：

```sh
git clone https://github.com/w31r4/open-office-artifact-tool.git
cd open-office-artifact-tool
npm install
node examples/create-xlsx-dashboard.mjs
```

正式发布后可直接使用：

```sh
npm install open-office-artifact-tool
```

下面创建一个工作簿并完成一次真实 XLSX 导出/导入：

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

## 原生 Skills

仓库采用四个插件包、五个 Skill 的参考兼容结构：

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) — 依赖宿主提供的实时 Excel 会话
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)

Office Skill 的普通文件工作流统一调用 OpenChestnut。PDF Skill 在 ReportLab、pdfplumber、pypdf、PyMuPDF、Poppler、pyHanko 和 veraPDF 之间显式选路，不做静默 fallback；缺少所需 Provider 时会明确失败。安装关系见 [PDF Provider Matrix](skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md)，具体适配进度见 [Reference Skill 兼容性](docs/reference-skills.md)。

## 必须知道的边界

- 导入 Office 文件时，未建模对象只有在模型仍携带原始包快照且相关拓扑未被不支持的操作改变时才能原样保留；否则导出会明确失败。
- 任意已有 PDF 不能像 Word 一样可靠地自动重排全文；现有文件编辑必须落在明确、可验证的有界操作中。
- PDF 签名与 LTV 交给 pyHanko，PDF/A 与 PDF/UA 机器验证交给 veraPDF；OCR 和复杂结构修复仍依赖外部工具。
- PyMuPDF/MuPDF 不随 MIT npm 包分发。使用者必须单独安装，并明确接受 GNU AGPL 或商业许可。
- LibreOffice、Poppler、Playwright 和原生 Office Bridge 是渲染/验证工具，不是隐藏的 Office codec fallback。

## 开发与验证

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

OpenChestnut 的确定性构建、C# 测试和协议生成命令见 [发布门禁](docs/release.md)。架构细节见 [运行时架构](docs/reference-runtime-architecture.md)，Agent 黑盒评测见 [PromptBench](docs/agent-evals.md)，完整 API 见 [API Reference](docs/api.md)。

## License

[MIT](LICENSE)。第三方运行时许可与来源见 `THIRD_PARTY_NOTICES.md` 和 OpenChestnut runtime notices。
