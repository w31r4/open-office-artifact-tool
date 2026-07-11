# open-office-artifact-tool 交接文档

生成时间：2026-07-11  
项目路径：`/Users/zfang/workspace/open-office-artifact-tool`  
远端仓库：`https://github.com/w31r4/open-office-artifact-tool.git`  
授权参考包：项目内 submodule `reference/office-artifact-tool`，来源 `/Users/zfang/workspace/office-artifact-tool`，包名 `office-artifact-tool@2.8.22`  
定位说明：`open-office-artifact-tool` 的明确目标是做一个**可开源、可发布 npm/GitHub 的 `office-artifact-tool` clean-room 替代实现**：基于公开标准、公开库和可合法使用的本地/云端运行时，复现参考包的可观察 agent-facing 能力，包括 Office/PDF artifact 创建、导入、检查、渲染和验证。

## 0. 授权参考包与使用边界

本项目的目标不是做 demo，也不是包装私有运行时，而是做出 `office-artifact-tool` 的**开源 clean-room 版本/替代实现**，对齐其可观察的 agent-facing 能力和使用体验，同时不复制其实现。

这里的 `reference-skills/` 也是项目目标的一部分：它们不是一次性附件，而是 agent 操作层的目标规格和工作流样本。最终形态应当让 agent 能用 Documents / Spreadsheets / Presentations / PDF 这几类 skill 的工作流，调用 `open-office-artifact-tool` 的公开 API，完整完成 Office/PDF artifact 的创建、编辑、导入导出、inspect/resolve、render/preview、verify/QA。接手者后续应把这些参考 skill 逐步 clean-room 改造成指向本项目实现的可运行 skill/fixtures/tests，而不是只把它们当文档阅读材料。

参考包位置：

- 项目内 submodule：`/Users/zfang/workspace/open-office-artifact-tool/reference/office-artifact-tool`
- Submodule remote：`https://github.com/w31r4/office-artifact-tool.git`
- 本地原始路径：`/Users/zfang/workspace/office-artifact-tool`
- 包名/版本：`office-artifact-tool@2.8.22`
- 入口：`dist/artifact_tool.mjs`
- JSX 子路径：`dist/presentation-jsx/`
- 参考包自带说明：`reference/office-artifact-tool/README.md`
- 参考包 package manifest：`reference/office-artifact-tool/package.json`

本交接目录同时内置了四个目标参考 skill 包的完整内容，便于接手者不用重新翻 本地参考资料：

- Documents：`reference-skills/documents/`，主说明 `reference-skills/documents/skills/documents/SKILL.md`
- Spreadsheets：`reference-skills/spreadsheets/`，主说明 `reference-skills/spreadsheets/skills/spreadsheets/SKILL.md`，另含 `skills/excel-live-control/`
- Presentations：`reference-skills/presentations/`，主说明 `reference-skills/presentations/skills/presentations/SKILL.md`
- PDF：`reference-skills/pdf/`，主说明 `reference-skills/pdf/skills/pdf/SKILL.md`

原始来源是 本地参考资料：`local reference source`。

接手者可以使用参考包来观察公开 API、help/inspect/render/export 行为、测试输入输出、包结构和能力边界；但不得把参考包的 reference implementation internals复制进 `open-office-artifact-tool`。当前项目实现必须继续基于公开标准、公开库、OOXML/PDF 规范、OpenXML SDK、Microsoft Office native automation、Playwright、LibreOffice、Poppler、PDF.js、sharp/canvas 等可合法使用的技术独立实现。

> 敏感信息处理：本文未包含 API key、密码、token、个人身份信息或私有凭据。npm 发布状态只记录为“未认证/未发布”，不包含任何账号细节。

## 1. 当前仓库状态

- 当前分支：`main`
- 当前 worktree：干净，无未提交改动
- 与远端关系：`main...origin/main`，当前未显示 ahead/behind
- 最新提交：`a6feeb1 Add COUNTIFS SUMIFS formulas`
- 最近提交和完整变更请直接看：
  - `git log --oneline`
  - `git show <commit>`
  - `docs/coverage.md`
  - `docs/api.md`
  - `docs/release.md`

## 2. 已完成的主要工作

不要把下面当作完整 diff；完整事实以仓库文件、提交历史和测试为准。这里仅做交接级别概览。

### 2.1 统一 JS/TS agent-facing facade

已保留 ESM npm 包入口，并围绕四类 artifact 建立了基本一致的 API 风格：

- Spreadsheet：`Workbook` / `SpreadsheetFile`
- Document：`DocumentModel` / `DocumentFile`
- Presentation：`Presentation` / `PresentationFile`
- PDF：`PdfArtifact` / `PdfFile`
- 共享能力：`FileBlob`、`inspect`、`resolve`、`render`/`preview`、`verify`/QA、`help`、生成 API 文档

关键文件：

- `src/index.mjs`
- `docs/api.md`
- `scripts/generate-api-docs.mjs`
- `test/help.mjs`
- `test/verify.mjs`
- `test/render.mjs`

### 2.2 Spreadsheets / XLSX

已实现并测试的能力包括：

- worksheet/range/value/formula 基础模型
- 公式依赖图、cycle/missing-sheet 报告、trace/inspect/verify 集成
- 结构化引用、defined names、shared formula、array formula / dynamic spill hints
- 公式目录已覆盖常见聚合、逻辑、文本、lookup、条件聚合和动态数组函数，例如：`SUM`、`AVERAGE`、`COUNTIF(S)`、`SUMIF(S)`、`INDEX`、`MATCH`、`VLOOKUP`、`XLOOKUP`、`SEQUENCE`、`FILTER`、`UNIQUE`、`SORT` 等
- native OOXML XLSX 导出/导入覆盖 shared strings、styles、tables、charts、images/drawings、sparklines、threaded comments/persons、data validations、conditional formatting、pivot cache records 等
- conditional formatting 已含 color scale 计算与 roundtrip
- style 支持已扩展到 alignment / borders

参考文件：

- `test/spreadsheet.mjs`
- `docs/coverage.md` 的 Spreadsheets 部分
- 最近相关提交见 `git log --oneline`，例如 `a6feeb1`、`4a1aef2`、`9b72d4e`、`79bfd82`、`2ec3e53`

### 2.3 Documents / DOCX

已实现并测试的能力包括：

- styled paragraphs、run-level style spans、lists、tables、headers/footers、sections、images、hyperlinks、fields、citations、comments、tracked insertions/deletions
- DOCX export/import 写入真实 WordprocessingML 包结构，而不是只靠自定义 metadata
- styles.xml / numbering.xml / comments.xml / header/footer / relationships / image media / section setup / tracked changes 等基础 roundtrip
- document verify 已覆盖 fake list、invalid links/citations、missing styles、table geometry、image metadata、section setup、dangling comments、tracked-change sanity 等方向
- SVG/page-layout preview 以及通过 renderer/native bridge 的 render gate 入口已存在

参考文件：

- `test/document.mjs`
- `docs/coverage.md` 的 Documents 部分

### 2.4 Presentations / PPTX

已实现并测试的能力包括：

- slides、shapes、textboxes、tables、charts、images、connectors、notes、comments/threads、theme/layout facades
- compose/JSX layout、autoLayout、layout JSON、layout QA
- PPTX export/import 已恢复 clean-room generated PPTX 中的 native text/table/chart/image/connector/notes/comments/theme/layout 信息
- chart schema 已扩展到 bar/line/pie 等基础类型，并支持 title/categories/series/colors/axis/legend/data label metadata
- verify 覆盖 overlap、off-canvas、text/table overflow、connector endpoint、chart/data consistency、image validity、placeholder/template gaps 等

参考文件：

- `test/presentation.mjs`
- `test/presentation-jsx.mjs`
- `docs/coverage.md` 的 Presentations 部分

### 2.5 PDF

已实现并测试的能力包括：

- modeled PDF pages、text、positioned text items、tables、images、chart regions、layout regions
- export/import metadata roundtrip
- arbitrary PDF parsing adapter 入口，含 optional PDF.js adapter 和 fallback heuristic parser
- positioned text geometry 的 table reconstruction
- SVG preview、layout JSON、render adapter 入口
- verify 覆盖 empty pages、unicode dash、table/image bounds、text extraction sanity、page geometry 等方向

参考文件：

- `test/pdf.mjs`
- `examples/parse-render-pdf.mjs`
- `docs/coverage.md` 的 PDF 部分

### 2.6 Renderer architecture

已完成 pluggable renderer adapter 架构：

- Playwright SVG/HTML → PNG/WebP/JPEG/PDF
- sharp adapter
- node-canvas adapter
- Poppler adapter
- LibreOffice adapter
- native Office adapter via Node/C# bridge

重要行为：缺少 optional dependency / CLI 时会明确报错或在测试中跳过，不会假装 raster 支持已经可用。

参考文件：

- `src/renderers/playwright.mjs`
- `src/renderers/sharp.mjs`
- `src/renderers/canvas.mjs`
- `src/renderers/poppler.mjs`
- `src/renderers/libreoffice.mjs`
- `test/renderer-adapters.mjs`
- `test/playwright-renderer.mjs`

### 2.7 Native Office / C# bridge

已新增 optional native bridge，不让核心 npm 包强依赖 Windows 或 Microsoft Office：

- Node wrapper：`src/native/office-bridge.mjs`
- C#/.NET sidecar：`native/OfficeBridge/`
- JSON stdin/stdout protocol
- timeout、temp directory isolation、cleanup、structured errors
- Windows + Office 环境下设计支持 DOCX/XLSX/PPTX open/save/export/recalculate/update fields/render 等操作
- 测试在无 Office/无 dotnet 环境下 graceful skip

参考文件：

- `native/OfficeBridge/src/Program.cs`
- `native/OfficeBridge/src/BridgeProtocol.cs`
- `native/OfficeBridge/src/OfficeAutomation.cs`
- `native/OfficeBridge/tests/BridgeProtocolTests.cs`
- `test/office-bridge.mjs`

### 2.8 Docs / examples / CI / release readiness

已存在：

- README：项目定位、四类 artifact、renderer adapters、Playwright、native bridge、examples、release check
- `docs/coverage.md`：覆盖矩阵，当前保持 partial/done/todo，不虚标全部完成
- `docs/api.md`：由 `HELP_CATALOG` 生成
- examples：DOCX report、XLSX dashboard、PPTX compose、PDF parse/render、Playwright render、native Office render
- CI：基础 npm test/docs/pack，不依赖 Microsoft Office；dotnet bridge tests 条件执行
- release workflow：manual dry-run / npm publish / tag / GitHub release，受 npm token/auth 控制

参考文件：

- `README.md`
- `docs/coverage.md`
- `docs/api.md`
- `docs/release.md`
- `.github/workflows/ci.yml`
- `.github/workflows/release.yml`
- `examples/`

## 3. 当前测试 / 发布状态

最近记录的本地 gate 状态：

- `npm test`：通过
- `npm run test:pack`：通过
- `npm run docs:api`：通过
- Playwright renderer smoke：本机未安装 Playwright/Chromium 时按设计跳过
- `dotnet test native/OfficeBridge`：本机无 `dotnet` 时无法运行，CI/有 dotnet 的机器可跑
- npm publish：未完成，当前阻塞是 npm auth 不可用；详见 `docs/release.md`

接手者开始工作前建议重新执行：

```bash
cd /Users/zfang/workspace/open-office-artifact-tool
npm test
npm run docs:api
npm run test:pack
```

如本机具备对应环境，再执行：

```bash
dotnet test native/OfficeBridge
# 若已安装 Playwright/Chromium，也跑 renderer smoke tests
```

## 4. 还没做完 / 明确阻塞项

这些不是“bug 全部没做”，而是离最终发布级完整套件还差的主要 gap。

### 4.1 外部运行时实机验证

- 还缺 Windows + Microsoft Office 环境下 native bridge 的真实集成验证
- 还缺 LibreOffice / Poppler / Playwright 在完整安装环境下的跨平台验证矩阵
- 现有本机环境缺 `dotnet`，Playwright/Chromium 也未安装，所以相关测试只能 graceful skip

### 4.2 npm / GitHub release

- npm auth 当前不可用，包尚未真正发布到 npm
- tag/release 流程已经有 workflow，但还需要有权限的人执行并验证
- 发布前必须重跑 release gate，并确认 package 内容和 docs 都是 publish-ready

### 4.3 Spreadsheet 仍需增强

优先级较高的缺口：

- 扩展公式目录：建议下一批补 `SUMPRODUCT`、`HLOOKUP`、`IFERROR`、`ISNUMBER`、`ISTEXT`、`ISBLANK`、`ISERROR`、`AVERAGEIF`、`AVERAGEIFS`
- 更接近 Excel 的 coercion/error semantics
- 更完整 structured-reference escaping / intersections
- richer native pivot refresh/interoperability
- richer conditional formatting compatibility
- workbook render-backed QA 更完整
- 可考虑补 `freezePanes` / `autoFilter` native XLSX export/import

### 4.4 Document 仍需增强

- DOCX table cell style / borders / shading fidelity
- complex-script / theme / run import fidelity
- 更真实的 pagination/text measurement
- 更高层 OpenXML patch recipes
- 实机 Word/LibreOffice render evidence

### 4.5 Presentation 仍需增强

- richer chart types：例如 area、scatter、doughnut
- master/theme/layout fidelity
- comments interoperability
- placeholder/template fidelity gates
- render-backed visual QA
- 更广泛 third-party PPTX import fidelity

### 4.6 PDF 仍需增强

- arbitrary PDF parsing fidelity
- multi-line / spanning-cell table reconstruction
- byte-level image extraction
- Poppler/PDFium alternatives 或更强 PDF.js integration
- raster visual diff workflow
- 更完整 pagination/report layout/typography/chart styling

### 4.7 Help/API schema 仍需增强

- `HELP_CATALOG` 已有 selected high-traffic APIs 的 schema metadata
- 但还不是全量覆盖
- 每新增 API 或能力，都要同步补 help entry、docs generation、smoke tests 和 `docs/coverage.md`

## 5. 推荐下一步执行顺序

建议下一位接手者按这个节奏走，避免范围失控：

1. **进入仓库后先确认状态**

   ```bash
   cd /Users/zfang/workspace/open-office-artifact-tool
   git status -sb
   git log --oneline -8
   ```

2. **先读覆盖矩阵，不凭印象选任务**

   - `docs/coverage.md`
   - `docs/release.md`
   - 必要时看 `docs/api.md`

3. **选一个小而高价值的 gap 做完整闭环**

   推荐第一优先级：扩展 Spreadsheet formula catalog。原因：

   - 对目标验收里的 “broad formula catalog” 直接加分
   - 范围相对清晰
   - 容易写单元测试、coverage 更新和 API/docs 更新
   - 不依赖外部 Office/Playwright/dotnet 环境

   推荐第一批函数：

   - `SUMPRODUCT`
   - `HLOOKUP`
   - `IFERROR`
   - `ISNUMBER`
   - `ISTEXT`
   - `ISBLANK`
   - `ISERROR`
   - `AVERAGEIF`
   - `AVERAGEIFS`

4. **每个功能必须闭环**

   每次改动至少包含：

   - 实现：通常在 `src/index.mjs`
   - 测试：优先在对应 `test/*.mjs`
   - 文档/coverage：`docs/coverage.md`，必要时更新 README/API docs source
   - 重新生成 API docs：`npm run docs:api`
   - 跑 gate：`npm test`、`npm run test:pack`

5. **提交与推送**

   用户原始要求是每阶段完成后 commit 并 push 到 `main`。提交信息必须带：

   ```text
   Co-Authored-By: Enter Code <noreply@enter.pro>
   ```

   不要跳过 hooks，不要 force push，不要用 destructive git command。

## 6. 对外表述注意事项

接手者在 README、release note、PR description、handoff 或对外沟通中，应统一使用以下表述：

- 这是一个独立、已获授权的 clean-room Office/PDF artifact toolkit 项目
- 实现基于公开标准、公开规范、公开库和可合法使用的本地运行时
- 不包含私有源码、私有 bundle、私有 runtime module 或未授权实现
- 能力目标是提供面向 agent 的 Office/PDF artifact 创建、导入、检查、渲染、验证和发布级工具链

发布前建议额外做一次文案审查：把所有外部可见文档统一成上述定位，避免遗留的比较式或不准确表述。

## 7. Suggested skills

下一位 agent 正常编码时不一定需要额外 skill；优先直接读仓库、改代码、跑测试。以下 skill 只在对应场景调用：

- `handoff`：如果继续交接给下一位开发者，生成新的临时目录交接文档。
- `plain-mr-summary`：如果需要给维护者/评审者用人话解释某个 commit series、branch、PR 或 diff。
- `retrospective-doc`：如果用户要求把这轮长期建设整理成正式复盘或项目总结。
- `stuck`：如果在 native Office、Playwright、dotnet、npm publish 等环境问题上多次尝试仍卡住。
- `ultrareview`：仅当用户显式要求 `/ultrareview` 时，用于对变更做多轮复审。

## 8. 交接结论

当前项目不是空壳 demo：四类 artifact 的核心 facade、基础 OOXML/PDF roundtrip、inspect/resolve/help/render/verify、renderer adapters、native bridge、docs、examples、CI/release skeleton 都已经搭起来，并且最近主线保持 clean。

但它也还不是“完全发布结束”的状态：主要差距集中在外部运行时实机验证、npm 发布权限、Office/PDF 高保真细节、完整 API schema 覆盖和更广泛兼容性测试。下一位接手者应从 `docs/coverage.md` 出发，每次只选一个高价值 gap，做到实现 + 测试 + 文档 + gate + commit/push 的完整闭环。
