# Pek.Office - 办公自动化库

## 项目简介

`Pek.Office` 是一个功能全面的 .NET 办公自动化库，覆盖 **Excel（xlsx/xls）、Word（docx/doc）、PPT（pptx/ppt）、PDF**、Markdown、RTF、ODS、Email（eml/msg）、iCalendar、vCard、EPUB、XPS 以及基于 OLE2/BIFF8 的老式二进制格式，致力于为 .NET 社区提供零依赖、MIT 许可的高质量办公自动化能力。

- 仅依赖 `NewLife.Core`，零外部依赖
- 支持 `net45/netstandard2.0/netstandard2.1`
- 直接基于 Open XML（ZIP + XML）/ PDF / OLE2 文件结构实现
- 统一命名空间 `Pek.Office`，API 风格一致，学习成本低

> CSV 读写能力由 `NewLife.Core` 的 `CsvFile` 提供，本仓库不重复实现 CSV。

## 已实现能力（Excel）

### ExcelWriter（写入）

- 多工作表写入、表头与多行数据写入
- 单元格样式：字体/背景/边框/对齐/换行/数字格式
- 布局能力：列宽、行高、合并单元格、冻结窗格、自动筛选
- 高级能力：超链接、下拉数据验证、条件格式、插入图片、图表（柱/折/饼/散点）、数据透视表
- 页面能力：页面方向、纸张大小、页边距、页眉页脚、打印标题行
- 安全能力：工作表保护
- 便捷能力：`WriteObjects<T>()`、`WriteDataTable()`

### ExcelReader（读取）

- 读取多工作表行数据（自动类型转换）
- `ReadObjects<T>()` 对象映射导入（支持 `DisplayName/Description`）
- `ReadDataTable()` 导入为 `DataTable`
- `GetMergeRanges()` 读取合并区域

### ExcelTemplate（模板填充）

- 基于现有 xlsx 模板执行 `{{Key}}` 占位符替换
- 支持输出到文件或流

### BiffReader（xls 安全读取）

- 基于 OLE2/CFB 容器，解析 BIFF8 格式工作表流
- `ReadSheet()` 读取工作表行数据
- `ReadObjects<T>()` 对象映射导入
- 支持字符串、数字、布尔、日期、RK 编码等多种单元格类型

## 已实现能力（Word）

### WordWriter / WordReader / WordTemplate

- docx 全面读写：段落、表格、图片、标题、列表、超链接、页眉页脚、页面设置
- 模板填充：`{{Key}}` 占位符替换，保留原样式
- `WordHtmlConverter`：docx 转语义化 HTML（段落→`<p>`，表格→`<table>`，标题→`<h1>`-`<h6>`）
- `WordPdfConverter`：docx 转 PDF（内容映射型，纳入 PdfFluentDocument）
- `DocReader`：读取 doc 格式（OLE2 容器，提取纯文本和段落结构）

## 已实现能力（PPT）

### PptxWriter / PptxReader / PptxTemplate

- pptx 全面读写：幻灯片管理、文本框、表格、图片、基本形状、超链接、备注、模板填充
- 布局设计：母版、版式、背景、尺寸、幻灯片切换动画
- `PptReader`：读取 ppt 格式（OLE2 容器，提取幻灯片文本）

## 已实现能力（PDF）

### PdfWriter / PdfReader / PdfDocument

- 创建 PDF：文本排版、表格、图片、页眉页脚、中文字体支持（Type0/STSong-Light，UniGB-UCS2-H 编码）
- `PdfFluentDocument`：声明式布局 API（Row/Column/Text/Image/Table），自动分页
- PDF 操作：合并/拆分、水印、页面旋转、文字覆盖、可点击链接、书签导航
- `PdfEncryptor`：加密/权限控制
- 文本提取：逐页/全文，带坐标信息提取

## 已实现能力（Markdown / RTF / ODS）

### Markdown

- `MarkdownParser`：解析 CommonMark + GFM（标题/段落/代码块/列表/引用/粗斜体/图片/表格/任务列表）
- `MarkdownWriter`：将 `MarkdownDocument` 序列化为 .md 文本（包括 GFM 扩展）
- `MarkdownHtmlConverter`：Markdown 转换为语义化 HTML
- `MarkdownWordConverter`：Markdown → Word（docx），标题/段落/列表/表格映射为 Word 样式
- `MarkdownPdfConverter`：Markdown → PDF（通过 PdfFluentDocument 直出 PDF）
- `MarkdownDocument.ToWord()` / `ToPdf()` 快捷方法

### RTF

- `RtfReader`：解析 RTF 文件，提取段落（字体/颜色/对齐）、表格、文档属性、嵌入图片（PNG/JPEG/WMF/EMF）
- `RtfWriter`：创建 RTF 文件（段落格式/表格/模板填充）；`AddImage()` 支持嵌入 PNG/JPEG 图片（十六进制编码）

### ODS

- `OdsReader`：读取 LibreOffice ODS 格式，提取多工作表单元格数据、合并单元格、单元格样式（字体/颜色/背景/对齐）
- `OdsWriter`：生成 ODS 文件（多工作表、字符串/数字/日期/布尔/公式、单元格样式写入）

## OLE2/CFB 容器

- `CfbReader`：解析 Compound File Binary 容器（安全寻找 Stream，支持 v3/v4 水印）
- `CfbWriter`：生成合法 CFB 文件（写入流和目录树）
- 支撟 xls（Excel 97-2003）、doc（Word 97-2003）、ppt（PPT 97-2003）共用基础层

## 快速开始

### 统一文本/Markdown 提取

```csharp
using NewLife.Office;

// 一行代码提取任意格式纯文本
String? text = OfficeFactory.ReadText("document.docx");
String? text2 = OfficeFactory.ReadText("report.xlsx");

// 提取为 Markdown 格式（保留结构信息）
String? md = OfficeFactory.ReadMarkdown("slides.pptx");

// 从流中提取
using var fs = File.OpenRead("data.pdf");
String? pdfText = OfficeFactory.ReadText(fs, ".pdf");
```

也可使用接口直接操作：

```csharp
using var reader = new WordReader("document.docx");
String? text = (reader as ITextExtractable)?.ExtractText();
String? md = (reader as IMarkdownExtractable)?.ExtractMarkdown();
```

### Excel 读写

```csharp
using NewLife.Office;

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);

writer.WriteHeader("Data", new[] { "编号", "姓名", "金额" }, CellStyle.Header);
writer.WriteRow("Data", new Object?[] { 1, "张三", 1234.56m }, new CellStyle
{
    Border = CellBorderStyle.Thin,
    NumberFormat = "#,##0.00"
});

writer.FreezePane("Data", 1);
writer.SetAutoFilter("Data", "A1:C1");
writer.Save();

ms.Position = 0;
var reader = new ExcelReader(ms, Encoding.UTF8);
var rows = reader.ReadRows("Data").ToList();
```

### 对象集合导出

```csharp
public class UserInfo
{
    [DisplayName("编号")]
    public Int32 Id { get; set; }

    [DisplayName("姓名")]
    public String Name { get; set; }

    [DisplayName("金额")]
    public Decimal Amount { get; set; }
}

var list = new List<UserInfo>
{
    new() { Id = 1, Name = "张三", Amount = 1234.56m },
    new() { Id = 2, Name = "李四", Amount = 9876.00m },
};

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);
writer.WriteObjects("用户列表", list);
writer.Save();
```

### 对象集合导入

```csharp
using var fs = File.OpenRead("users.xlsx");
var reader = new ExcelReader(fs);
var users = reader.ReadObjects<UserInfo>("用户列表");
```

### 模板填充

```csharp
// 基于 xlsx 模板替换 {{Key}} 占位符，保留原有样式
var template = new ExcelTemplate("report_template.xlsx");
template.Fill(new Dictionary<String, Object>
{
    ["Title"] = "2025年度报告",
    ["Author"] = "张三",
    ["Date"] = DateTime.Today.ToString("yyyy-MM-dd"),
});
template.Save("report_output.xlsx");
```

### 单元格样式

```csharp
var style = new CellStyle
{
    Bold = true,
    FontColor = "FFFFFF",        // 白色字体
    BackgroundColor = "4472C4",  // 蓝色背景
    HAlign = HorizontalAlignment.Center,
    Border = CellBorderStyle.Thin,
    NumberFormat = "#,##0.00",
};

writer.WriteHeader("Sheet1", new[] { "编号", "姓名", "金额" }, style);
writer.MergeCell("Sheet1", "A1:C1");  // 合并标题行
writer.SetRowHeight("Sheet1", 1, 30); // 设置行高
```

## 已实现能力（邮件 / 日历 / 联系人 / 电子书）

### EML（邮件）

- `EmlReader`：解析 RFC 5322 + MIME 邮件，含 multipart/mixed、multipart/alternative
- `EmlWriter`：生成合法 MIME 邮件（文本/HTML 正文、附件、内嵌图片）
- 支持 RFC 2047 Base64/Q-Encoding 头部解码、quoted-printable 正文解码

### MSG（Outlook 邮件）

- `MsgReader`：基于 OLE2/CFB 容器解析 Outlook .msg 文件，提取 MAPI 属性
- 支持主题（PR_SUBJECT）、发件人、收件人列表（To/Cc/Bcc）
- 支持纯文本正文（PR_BODY）和 HTML 正文（PR_HTML_BODY）
- 支持多附件提取（`__attach_version1.0_#NN` 子存储，含文件名和二进制数据）

### iCalendar（日历）

- `ICalReader`：解析 RFC 5545 .ics 文件，提取 VEVENT/VTODO/VJOURNAL 及重复规则
- `ICalWriter`：生成符合 RFC 5545 的 .ics 文件，支持 RRULE 序列化和 75 字节行折叠

### vCard（联系人）

- `VCardReader`：解析 RFC 6350 .vcf 文件，支持 vCard 3.0/4.0，含 FN/N/ORG/TEL/EMAIL/ADR/BDAY/PHOTO
- `VCardWriter`：生成 vCard 4.0 格式，支持多联系人批量写入

### EPUB（电子书）

- `EpubReader`：读取 EPUB 2/3 文件（ZIP + OPF + XHTML），提取元数据和章节内容
- `EpubWriter`：生成符合 EPUB 3 规范的电子书（mimetype/container.xml/OPF/Nav/章节 XHTML/CSS/封面图片）

## 已实现能力（XPS）

### XpsReader / XpsWriter

- `XpsReader`：解析 XPS（XML Paper Specification）文件（ZIP + XML），遍历 FixedDocumentSequence → FixedDocument → FixedPage
- 文本提取：从 `<Glyphs UnicodeString="...">` 元素提取文本内容
- 图片提取：`ExtractImages()` 返回所有嵌入 PNG/JPEG 资源
- 页面信息：`Width`/`Height` 从 `<FixedPage>` 属性读取
- 元数据：`ReadProperties()` 解析 `docProps/core.xml`（标题/作者/主题）
- `XpsWriter`：生成合法 XPS 文件（`[Content_Types].xml` + `_rels/.rels` + FDS + .fpage）
- `AddPage(width, height, IEnumerable<(text, x, y, fontSize)>)` 多页写入
- `SetProperties(XpsProperties)` 写入文档属性
- `AddImage(name, data, mime)` 嵌入图片资源
- `Save(path|stream)` / `ToBytes()` 输出

## 测试状态

- 当前测试集：`395` 项
- 通过：`395` 项
- 失败：`0` 项

## 已实现格式概览

| 格式 | 读取 | 写入 | 模板 | 转换 | 说明 |
|------|:---:|:---:|:---:|:---:|------|
| **Excel xlsx** | ✅ | ✅ | ✅ | — | 含样式/图表/透视表/条件格式/图片 |
| **Excel xls** | ✅ | 🔄规划 | — | — | OLE2/BIFF8，BiffReader 已实现读取 |
| **Word docx** | ✅ | ✅ | ✅ | HTML/PDF | 含段落/表格/图片/页眉页脚 |
| **Word doc** | ✅ | — | — | — | MS-DOC，DocReader 提取纯文本 |
| **PPT pptx** | ✅ | ✅ | ✅ | 🔄规划 | 含文本框/图片/表格/图表/母版 |
| **PPT ppt** | ✅ | — | — | — | MS-PPT，PptReader 提取幻灯片文本 |
| **PDF** | ✅ | ✅ | — | — | 含 Fluent API/合并/拆分/加密 |
| **Markdown** | ✅ | ✅ | — | HTML/Word/PDF | CommonMark + GFM，含 AST 模型，→Word/→PDF 已实现 |
| **RTF** | ✅ | ✅ | ✅ | — | 段落/表格/模板填充/图片嵌入（PNG/JPEG） |
| **ODS** | ✅ | ✅ | — | — | LibreOffice 原生格式 |
| **EML** | ✅ | ✅ | — | — | RFC 5322 + MIME，附件/内嵌图片 |
| **MSG** | ✅ | — | — | — | Outlook，OLE2 + MAPI，提取主题/正文/附件 |
| **iCalendar** | ✅ | ✅ | — | — | RFC 5545 .ics，VEVENT/VTODO/RRULE |
| **vCard** | ✅ | ✅ | — | — | RFC 6350 .vcf，vCard 3.0/4.0 |
| **EPUB** | ✅ | ✅ | — | — | EPUB 2/3，OPF/Nav/章节/封面 |
| **XPS** | ✅ | ✅ | — | — | ZIP+XML，Glyphs 文本提取/多页写入/图片/元数据 |

## 代码结构

```
NewLife.Office/
├── Excel/        # xlsx 读写：ExcelWriter / ExcelReader / ExcelTemplate / CellStyle / ExcelHelper / BiffReader
├── Word/         # docx 读写：WordWriter / WordReader / WordTemplate / DocReader / WordHtmlConverter / WordPdfConverter
├── PPT/          # pptx 读写：PptxWriter / PptxReader / PptxTemplate / PptReader
├── PDF/          # PDF 读写：PdfWriter / PdfReader / PdfDocument / PdfFluentDocument / PdfEncryptor
├── OLE2/         # OLE2/CFB 容器：CfbReader / CfbWriter
├── Markdown/     # Markdown 读写：MarkdownParser / MarkdownWriter / MarkdownHtmlConverter
├── Rtf/          # RTF 读写：RtfReader / RtfWriter
├── Ods/          # ODS 读写：OdsReader / OdsWriter
├── Mail/         # EML 邮件读写：EmlReader / EmlWriter / EmlMessage
├── Calendar/     # iCalendar .ics 读写：ICalReader / ICalWriter / ICalDocument
├── VCard/        # vCard .vcf 读写：VCardReader / VCardWriter / VCardContact
├── Epub/         # EPUB 读写：EpubReader / EpubWriter / EpubDocument
└── Xps/          # XPS 读写：XpsReader / XpsWriter
```

## 规划中

以下功能已在路线图中，但尚未实现：

### 写入 xls（BIFF8，Excel 97-2003）

- 生成合法 BIFF8 工作簿流并写入 OLE2 容器

### PPT → PDF / 图片

- pptx 每张幻灯片转换为 PDF 页面或 PNG/JPEG 图像

### ODS 合并单元格

- OdsReader 识别合并单元格（number-columns-spanned）