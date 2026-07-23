# OfficeKit

**给 Agent 的 Office 与 PDF 多功能工具箱。**

**简体中文** | [English](README.en.md)

OfficeKit 让 Agent 用同一套工作流创建、读取、编辑、检查、渲染和验证
DOCX、XLSX、PPTX 与 PDF。

它不是一个必须记住子命令的传统 CLI，而是三层组合：

- **Skills**：告诉 Agent 什么时候用什么能力、如何检查结果和何时拒绝操作。
- **JavaScript API**：提供对象模型、计算、Compose、inspect、render、verify 和显式 package patch。
- **原生引擎**：Office 使用 OpenChestnut C#/.NET WASM；PDF 使用懒加载的 MuPDF.js。

> 当前仓库包名仍为 `open-office-artifact-tool`，`OfficeKit` 是面向使用者的产品名。当前版本为 `0.3.0` release candidate，尚未正式发布到 npm。

## 30 秒部署

### 只装 Skills

不需要 clone 仓库，也不需要安装 .NET、Python 或 Office。把需要的 Skill
直接装进当前 Agent 的项目目录：

```sh
npx skills add w31r4/open-office-artifact-tool \
  --skill documents \
  --skill Spreadsheets \
  --skill Presentations \
  --skill pdf \
  --skill template-creator \
  --yes
```

需要全部原生 Skill 和开源模板时：

```sh
npx skills add w31r4/open-office-artifact-tool --skill '*' --yes
```

`npx skills` 会把 Skill 安装到 Agent 能识别的位置；不需要手工复制
`skills/` 目录。也可以加 `--global` 作为用户级安装，或用 `--agent`
明确指定宿主。

### 需要调用 JavaScript API

当前 npm 发布前，从 GitHub 安装；正式发布后把安装源替换成包名：

```sh
# 当前 release candidate
npm install github:w31r4/open-office-artifact-tool

# npm 正式发布后
npm install open-office-artifact-tool
```

消费者只需要 Node.js。npm 包自带 OpenChestnut 的 WASM runtime；不需要本机
安装 .NET。建议使用 Node.js 22 或更新版本。第一次执行 PDF 操作时才会初始化
MuPDF WASM；root import、普通 Office 操作和 npm install 都不会下载额外运行时。

安装后可以直接运行一个完整示例：

```sh
node node_modules/open-office-artifact-tool/examples/create-xlsx-dashboard.mjs
```

### 只试用一个 Skill

不想把 Skill 写入项目时，可生成一次性 Agent 提示：

```sh
npx skills use w31r4/open-office-artifact-tool --skill pdf
```

## 第一个 API 调用

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

更多可运行示例：

- [创建 DOCX 报告](examples/create-docx-report.mjs)
- [创建 XLSX 仪表盘](examples/create-xlsx-dashboard.mjs)
- [使用 Compose 创建 PPTX](examples/create-pptx-compose.mjs)
- [解析与渲染 PDF](examples/parse-render-pdf.mjs)

## 四类文件，一套 Agent 工作流

| 格式 | 默认引擎 | 适合做什么 |
| --- | --- | --- |
| DOCX | OpenChestnut C# WASM | 创建和编辑结构化文档、样式、分节、表格、图片、字段、评论和有界内容控件。 |
| XLSX | OpenChestnut C# WASM | 单元格、公式、样式、布局、表格、图片、验证、条件格式、评论、图表、sparklines、What-If 和有界 PivotTable。 |
| PPTX | OpenChestnut C# WASM | 形状、富文本、图片与可逆裁剪、表格、连接线、图表、备注、评论、Master/Layout 保真和有界源绑定编辑。 |
| PDF | PdfArtifact + MuPDF.js | 新建 Tagged PDF；读取、检查、渲染、表单、链接、批注、页面编辑、rewrite 脱敏和有界签名。 |

Office 的普通导入和导出只有一条 OpenChestnut 路径。无法安全建模的内容会绑定原始包并保留；不支持的修改会明确失败，不会悄悄切换到第二套 codec。

## PDF 专项能力按需加载

`mupdf@1.28.0` 是 PDF 的必需 npm 依赖，随正常 `npm install` 解析，运行时懒加载。
qpdf、Python、OCR、veraPDF/JRE 等大体积工具不进入 npm tarball，也不会通过
`postinstall` 或全局包管理器偷偷安装。

需要专项能力时，先让 Agent 通过公开 provider API 判断：

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

默认策略是 `disabled`。只有项目明确授权 `managed`，并满足平台、哈希、体积、
许可证和语言包限制时，Agent 才能 `ensure`；已有部署也可以明确选择
`system-only`。没有隐式 fallback。完整策略和安全边界见
[PDF Provider Setup](skills/pdf/skills/pdf/tasks/provider_setup.md)。

## Skills 与模板

仓库提供四个参考文件 Skill，以及模板创建和模板库能力：

- [Documents](skills/documents/skills/documents/SKILL.md)
- [Spreadsheets](skills/spreadsheets/skills/spreadsheets/SKILL.md)
- [Presentations](skills/presentations/skills/presentations/SKILL.md)
- [PDF](skills/pdf/skills/pdf/SKILL.md)
- [Template Creator](skills/template-creator/skills/template-creator/SKILL.md)
- [Office Template Library](skills/default-template-library/README.md) — 20 套 MIT 授权模板，仅仓库分发，不进入 npm runtime tarball。

模板不是另一套 codec。Agent 先把指定模板物化成新的输出文件，再用同一套
Office API 检查、编辑、渲染和验证；参考文件本身不会被覆盖。

## 交付前检查

推荐 Agent 在交付前固定执行：

```text
intent → inspect → resolve → edit/create → export → re-import → render → verify
```

核心 API 会尽可能返回可审计的结构化证据。源快照丢失、拓扑不可信、签名会失效、
PDF 操作可能残留旧 revision，或需要外部凭据时，操作会 fail closed。

## 开发与验证

```sh
npm test
npm run test:pack
npm run docs:api
npm run release:check
```

完整能力边界见 [coverage](docs/coverage.md)，API 见 [docs/api.md](docs/api.md)，
参考 Skill 兼容性见 [docs/reference-skills.md](docs/reference-skills.md)。

## 许可证

[GNU AGPL v3 或更高版本](LICENSE)。网络部署、修改和分发必须遵守 AGPL 的对应义务。
第三方运行时、MuPDF 和专项 provider 的许可证与来源见
`THIRD_PARTY_NOTICES.md` 及相关 runtime notices。
