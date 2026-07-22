# open-office-artifact-tool

**简体中文** | [English](README.en.md)

面向 Agent 的 Office 与 PDF 创建、读取、编辑、检查、渲染和验证工具箱。

`open-office-artifact-tool` 提供统一的 JavaScript 对象模型。DOCX、XLSX 和 PPTX 由仓库内的 **OpenChestnut**（C# + Open XML SDK + .NET WebAssembly）读写；PDF 使用独立语义模型与运行时懒加载的 **MuPDF.js** 原生管线。

> **当前状态：** `0.3.0` 发布候选。源码、可复现 WASM 和 npm tarball 已具备验证流程，但尚未执行正式 `npm publish`。

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

### PDF 运行时

官方 `mupdf@1.28.0` 是必需的 npm 依赖，会随正常的 `npm install` 一起解析安装；只有第一次读取、检查、渲染或编辑 PDF 时才初始化 WASM。项目没有 `postinstall`、额外下载器或全局环境写入。

qpdf、Python 专项工具、OCR、veraPDF/JRE 等由显式的 `open-office-artifact-tool/pdf/providers` 路由：先根据任务和检查结果返回 `ready`、`installable` 或 `blocked`，再由项目策略决定是否可安装。默认策略禁止下载；qpdf、Python foundation 和 Python specialists 都已有双平台、哈希固定且带证明的能力包。foundation 只提供隔离 CPython 上的 ReportLab、pdfplumber、pypdf 与 Pillow；specialists 提供 PyMuPDF、pikepdf、pyHanko 与证书校验器，依赖 qpdf 并需要显式接受 AGPL 或商业许可。OCR、Poppler、veraPDF/JRE 等仍会明确 `blocked`，不会偷偷降级到其它工具。已由部署方管理的运行时可使用显式 `system-only` 策略。详见 [PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md)。

## 为什么需要它

- **为 Agent 设计**：文件模型自带 `inspect`、`resolve`、`verify`、render 和 visual QA 原语。
- **保真优先**：无法安全建模的 Office 内容绑定原始包并原样保留；不支持的修改明确失败。
- **内置原生 Skills**：npm 随包提供 Documents、Spreadsheets、Presentations、PDF 和 Template Creator 五个插件包（含六个 Skill）；仓库另保留一个 MIT 授权、仅限仓库使用的 20 套 Office Template Library。需要宿主会话或外部 Provider 的工作流会明确说明前置条件。

## 支持范围

| 格式 | 文件管线 | 当前核心能力 |
| --- | --- | --- |
| XLSX | OpenChestnut C# WASM | 单元格与公式、样式与布局、表格、图片、基础验证、标准 data bar/icon set 等条件格式、评论、图表、sparklines、有界 What-If Data Tables 和有界原生 PivotTables。 |
| DOCX | OpenChestnut C# WASM | 结构化文本与样式、分节、页眉页脚、列表、表格、链接、字段、行内图片及有界浮动/环绕图片、经典评论、有界现代评论线程、无密码编辑限制，以及块级纯文本和行内/整单元格纯文本、规范复选框、规范下拉、可输入自定义值的规范组合框、严格 `YYYY-MM-DD` 日期选择器内容控件。 |
| PPTX | OpenChestnut C# WASM | 形状与富文本、图片及可逆裁剪、表格、连接线、图表、直接背景、纯文本演讲者备注、经典评论和有界 Office 2021 现代评论线程；Master/Layout 仅保真、不可编辑。 |
| PDF | 独立模型 + MuPDF.js | Tagged PDF 创建；任意 PDF 原生读取/检查/渲染；有界批注、表单、页面、元数据、链接和 rewrite/incremental 编辑；真实 rewrite 脱敏；有界本地 PKCS#12 签名与独立验签。专项工具经显式 provider 路由；严格 sanitize、PDF/UA、OCR 与高级签名仍要求独立证据。 |

完整且持续更新的边界见 [能力矩阵](https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/coverage.md)。

## 工作方式

```text
Agent / Skill
├─ Office → JavaScript model → OpenChestnut C# WASM → DOCX / XLSX / PPTX
├─ PDF    → PdfArtifact（新建）或 MuPDF.js（导入/编辑）→ PDF
└─ QA     → inspect / resolve → render → verify / visual QA
```

OpenChestnut 是普通 Office 导入/导出的唯一 parser/writer。显式 OOXML inspect/patch 只供高级用户手动调用，不会成为自动 fallback。

## 原生 Skills

仓库包含六个插件包、二十六个 Skill。其中前五个插件包的六个 Skill 随 npm 分发；最后一个模板库含二十个仅限仓库使用的模板 Skill：

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Excel Live Control](skills/spreadsheets/skills/excel-live-control/SKILL.md) — 依赖宿主提供实时 Excel 会话
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)
- [Template Creator](skills/template-creator/skills/template-creator/SKILL.md) — 从本地 DOCX、PPTX 或 XLSX 参考文件创建/显式更新可复用模板
- [Office Template Library](skills/default-template-library/README.md) — 20 套 MIT 授权的保留模板：7 个 DOCX、7 个 PPTX、6 个 XLSX；仅限仓库使用，不进入 npm tarball

前五个 `skills/<name>` 目录随包分发，具体加载方式由 Agent 宿主决定。Office Skill 的普通文件工作流统一调用 OpenChestnut。PDF Skill 默认通过随 npm 安装的 MuPDF.js 薄 CLI 调用同一组包 API；Template Creator 只在 `${OFFICE_ARTIFACT_HOME:-~/.office-artifact-tool}/skills` 下事务式保存用户明确提供的本地参考文件与 PNG 预览，不联网、不覆盖未点名模板。默认模板库保留来自 MIT 参考仓库的原始 Office 与 PNG 文件，并用哈希记录来源；Agent 必须先把指定模板物化为新的输出文件，绝不修改仓库内的参考文件。二十套模板都经过导入、无变更导出、二次导入与原生渲染验证；已验证的修改是受限的 PPTX 幻灯片名称、DOCX 更新域设置和 XLSX 普通文本单元格。复杂源绑定内容仍会明确失败，而不会静默重建或替换版式。Python 与系统工具只承担尚无等价实现的专项工作。职责见 [PDF Provider Matrix](skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md) 与 [模板来源边界](docs/template-library-provenance.md)。

## 必须知道的边界

- 要保留导入 Office 文件中的未建模对象，必须继续使用 import 返回的模型，并保持这些对象的结构不变；丢失源快照或修改不支持的拓扑时，导出失败。
- 任意已有 PDF 不能像 Word 一样可靠地自动重排全文；原文件编辑必须落在明确、可验证的有界操作中。
- 项目内 pyHanko、veraPDF、OCRmyPDF、pikepdf、qpdf 等适配器都是显式 provider 路线：仅当受管包已验证就绪，或部署方显式选择 `system-only` 运行时时才可执行。私钥、P12、HSM/远程签名凭据、TSA/LTV 和信任根永远由调用方提供；不会被下载器取得。
- 主动/辅助内容清理由有界 pikepdf 路线完成；它保留 metadata、表单值、XFA、批注和隐藏文字，不能当作完整 sanitize 或脱敏证明。
- MuPDF.js 能做有界原文件操作，但不能把任意 PDF 变成可自由重排的 Word 文档；rewrite 脱敏也不等于完整 sanitize。签名权限、残留、OCR 与 PDF/UA 仍需独立证据。
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

[GNU AGPL v3 或更高版本](LICENSE)。网络服务部署、修改和分发必须遵守 AGPL 的对应义务；第三方运行时许可与来源见 `THIRD_PARTY_NOTICES.md` 和 OpenChestnut runtime notices。
