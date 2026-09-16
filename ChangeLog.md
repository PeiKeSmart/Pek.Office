# Pek.Office 版本更新记录

> 本项目基于 [NewLife.Office](https://github.com/NewLifeX/NewLife.Office)（新生命开发团队）持续同步维护：
> 包名 / 程序集名为 `Pek.Office`（含 `Pek.Office.Rendering`、`Pek.Office.Cli`），命名空间保持 `NewLife.Office` 以兼容上游 API。
> 上游更早版本的完整更新记录请参见上游仓库。

## 2026-09-16

- **构建**：`Pek.Office.Cli`、`Test`、`XUnitTest`、`MarkdownBenchmark` 追加 `net11.0` 多目标（保留 `net10.0`，输出按框架分目录）；Readme 与 Doc 竞品分析的框架声明同步至 `net11.0`；修复 WordMailMergeTests 等 5 个测试类固定临时文件名在双框架并行运行时的争用冲突

## 2026-09-15（九月发版同步）

同步上游 `v1.4.2026.0902`，主要能力：

- **Word**：读取保真修复（RunProperties 三态、样式继承链）；日常编辑 API（FindText/ReplaceText、表格行列增删、GetCell/SetCellText）；多节 / 脚注尾注 / 文本框；doc→docx 转换（`OfficeFactory.ConvertDocToDocx`）
- **Excel**：xls 全面增强（BiffWriter 写入、样式 / 批注 / 命名范围 / 打印区域）；条件格式 dxf；图表类型扩展；切片器 / 线程化批注；`ReadRows` 流式读取
- **Markdown**：CommonMark 规范补全与解析性能多轮优化
- **PPT**：形状往返保真、图表 / 超链接增强、跨文件合并
- **PDF**：结构化表格提取、嵌入字体子集化
- **新增**：`Pek.Office.Cli` 命令行工具链（convert / doc2docx / info）、`OfficeFactory` 内容识别（Detect）、统一文本与 Markdown 提取（`ReadText` / `ReadMarkdown`）
- **构建**：主库与 Rendering 追加 `net11.0` 目标；补齐测试项目引用（CLI / Rendering / NPOI / OpenXml / SkiaSharp）；Rendering 去除不存在的签名文件引用

## 2026-05-16

- 新增 `ITextExtractable` / `IMarkdownExtractable` 统一文本提取接口（全格式覆盖：Excel/Word/PPT/PDF/RTF/ODS/EML/iCal/vCard/EPUB/XPS）
- `OfficeFactory` 新增 `ReadText()` / `ReadMarkdown()` 一行提取静态方法

## 2026-04-19

- 项目更名为 Pek.Office；依赖切换为 `DH.NCore`；目标框架恢复 net45/netstandard2.0/netstandard2.1；重写 Readme
