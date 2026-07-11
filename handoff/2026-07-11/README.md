# open-office-artifact-tool handoff bundle

生成时间：2026-07-11

这是给下一位开发者/agent 的完整交接目录，已正式放在 `open-office-artifact-tool` 项目目录内。

## 目录内容

- `HANDOFF.md`：中文交接主文档，说明当前项目状态、已完成、未完成、阻塞项、下一步建议。
- `reference-skills/`：项目内 agent 操作层的目标规格和工作流样本，来自 本地参考资料 的 Documents / Spreadsheets / Presentations / PDF 参考 skill 包完整内容；后续要 clean-room 改造成能调用本项目 API 的可运行 skills/fixtures/tests。
  - `reference-skills/documents/`
  - `reference-skills/spreadsheets/`
  - `reference-skills/presentations/`
  - `reference-skills/pdf/`
- `INDEX.md`：本目录的快速索引。

## 主要项目

- 开源实现目标：`/Users/zfang/workspace/open-office-artifact-tool`，目标是做一个可开源、可发布 npm/GitHub 的 `office-artifact-tool` clean-room 替代实现
- 授权运行时参考包：项目内 submodule `reference/office-artifact-tool`，来源 `/Users/zfang/workspace/office-artifact-tool` / `https://github.com/w31r4/office-artifact-tool.git`
- 本交接目录内置参考 skill 包：`reference-skills/*`

## 参考使用边界

这些参考 skill 包不仅用于观察目标能力、API 形态、工作流、QA gate、文档结构和测试/渲染习惯，也用于定义后续 agent 如何通过本项目完成 Office/PDF 套件编辑。后续工作应把它们 clean-room 改造成指向 `open-office-artifact-tool` 的项目内 skills/fixtures/tests。不要把 reference implementation internals复制进开源实现。

`open-office-artifact-tool` 仍必须用公开标准、公开库、OOXML/PDF 规范、OpenXML SDK、Microsoft Office native automation、Playwright、LibreOffice、Poppler、PDF.js、sharp/canvas 等可合法使用技术独立实现。

## 建议阅读顺序

1. `HANDOFF.md`
2. `reference-skills/documents/skills/documents/SKILL.md`
3. `reference-skills/spreadsheets/skills/spreadsheets/SKILL.md`
4. `reference-skills/presentations/skills/presentations/SKILL.md`
5. `reference-skills/pdf/skills/pdf/SKILL.md`
6. 回到项目仓库读：
   - `/Users/zfang/workspace/open-office-artifact-tool/docs/coverage.md`
   - `/Users/zfang/workspace/open-office-artifact-tool/docs/api.md`
   - `/Users/zfang/workspace/open-office-artifact-tool/docs/release.md`
