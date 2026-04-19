# Pek.Office

Pek.Office 是一个面向 .NET 的办公文档处理库，当前仓库覆盖 Excel、Word、PowerPoint、PDF、Markdown、RTF、ODS、EML、MSG、iCalendar、vCard、EPUB、XPS 等常见办公与文档格式。

项目当前 NuGet 包名为 Pek.Office。由于历史兼容原因，公开 API 仍主要位于 NewLife.Office 命名空间，README 中涉及该命名空间的示例均以当前源码为准。

## 项目定位

- 面向服务端或工具侧的文档读写、导入导出与格式转换
- 基于文件格式本身实现，不依赖 Office 或 LibreOffice 进程自动化
- 同时支持 OpenXML/ZIP+XML 与 OLE2/CFB 旧二进制格式
- 提供统一入口 OfficeFactory，用于识别支持的文档类型并创建读取器

## 目标框架与依赖

- 目标框架：net45、netstandard2.0、netstandard2.1
- 运行时核心依赖：DH.NCore
- 测试框架：xUnit

安装示例：

```bash
dotnet add package Pek.Office
```

## 当前实现模块

### Excel

- ExcelReader：读取 xlsx 工作簿、工作表行数据、对象集合、DataTable
- ExcelWriter：写入 xlsx，支持表头、数据行、样式、冻结窗格、自动筛选、数据验证、图片、页面设置、条件格式等常见能力
- ExcelTemplate：基于占位符执行模板填充
- BiffReader：读取 Excel 97-2003 的 xls 二进制文件
- BiffWriter：生成 Excel 97-2003 的 xls 文件，支持多工作表、对象集合和 DataTable 导出

### Word

- WordReader / WordWriter：读写 docx
- WordTemplate：docx 模板占位符填充
- WordHtmlConverter：docx 转 HTML
- WordPdfConverter：docx 转 PDF，属于内容映射型转换，不依赖外部渲染引擎
- DocReader：读取 Word 97-2003 的 doc 文件，可提取全文、段落和表格

### PowerPoint

- PptxReader / PptxWriter：读写 pptx
- PptxTemplate：pptx 模板填充
- PptReader：读取 PowerPoint 97-2003 的 ppt 文件，提取幻灯片文本
- PptxPdfConverter：pptx 转 PDF，当前按文本内容和版面顺序映射，不做像素级渲染

### PDF

- PdfWriter / PdfReader / PdfDocument：基础 PDF 读写能力
- PdfFluentDocument：声明式 PDF 排版接口
- PdfEncryptor：PDF 加密与权限处理

### Markdown

- MarkdownParser / MarkdownWriter：Markdown 解析与输出
- MarkdownHtmlConverter：Markdown 转 HTML
- MarkdownWordConverter：Markdown 转 Word
- MarkdownPdfConverter：Markdown 转 PDF
- MarkdownDocument：文档对象模型与快捷转换入口

### RTF

- RtfDocument / RtfReader / RtfWriter：RTF 解析与写入
- 支持段落、表格、颜色、字体、图片提取与图片嵌入
- 支持基于 {{Key}} 的模板填充

### ODS

- OdsReader / OdsWriter：读写 OpenDocument Spreadsheet
- 支持多工作表、对象映射、DataTable、基础单元格样式、合并区域读取

### 邮件与个人信息格式

- EmlReader / EmlWriter / EmlMessage：EML 邮件读写
- MsgReader：Outlook MSG 读取
- ICalReader / ICalWriter / ICalDocument：iCalendar 读写
- VCardReader / VCardWriter / VCardContact：vCard 读写

### 电子书与固定版式文档

- EpubReader / EpubWriter / EpubDocument：EPUB 2/3 读写
- XpsReader / XpsWriter：XPS 读写、文本提取、图片提取、元数据读写

### 基础容器与工厂

- CfbDocument / CfbReader：OLE2/CFB 容器解析与写入
- OfficeFactory：统一识别支持的扩展名并创建对应读取器

## OfficeFactory 支持格式

当前 OfficeFactory 内置支持以下 15 种扩展名：

- .xlsx
- .xls
- .docx
- .doc
- .pptx
- .ppt
- .pdf
- .rtf
- .ods
- .epub
- .vcf
- .eml
- .ics
- .md
- .xps

说明：

- 对于 xlsx、xls、docx、doc、pptx、ppt、pdf，会返回对应 Reader 实例
- 对于 rtf、ods、epub、md，会直接返回解析后的文档对象或结果对象
- 对于 eml、ics、vcf、xps，会返回对应读取器实例，由调用方继续读取文件内容

## 快速开始

### 读取任意支持的文件

```csharp
using NewLife.Office;

if (OfficeFactory.IsSupported(".xlsx"))
{
    using var reader = (ExcelReader)OfficeFactory.CreateReader("demo.xlsx");
    var rows = reader.ReadRows("Sheet1").ToList();
}
```

### 写入 xlsx

```csharp
using System.Text;
using NewLife.Office;

using var ms = new MemoryStream();
var writer = new ExcelWriter(ms);

writer.WriteHeader("Data", ["编号", "姓名", "金额"], CellStyle.Header);
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

### 写入 xls

```csharp
using NewLife.Office;

using var writer = new BiffWriter();
writer.WriteHeader(["姓名", "年龄", "成绩"]);
writer.WriteRow(new Object?[] { "Alice", 28, 95.5 });
writer.WriteRow(new Object?[] { "Bob", 31, 88.0 });
writer.Save("data.xls");
```

### docx 转 PDF

```csharp
using NewLife.Office;

var converter = new WordPdfConverter();
converter.ConvertToFile("input.docx", "output.pdf");
```

### pptx 转 PDF

```csharp
using NewLife.Office;

var converter = new PptxPdfConverter
{
    DocumentTitle = "演示文稿导出"
};
converter.Convert("slides.pptx", "slides.pdf");
```

## 能力概览

| 格式 | 读取 | 写入 | 模板 | 转换 | 说明 |
|------|:---:|:---:|:---:|:---:|------|
| Excel xlsx | ✅ | ✅ | ✅ | — | ExcelReader / ExcelWriter / ExcelTemplate |
| Excel xls | ✅ | ✅ | — | — | BiffReader / BiffWriter |
| Word docx | ✅ | ✅ | ✅ | HTML、PDF | WordReader / WordWriter / WordTemplate |
| Word doc | ✅ | — | — | — | DocReader |
| PowerPoint pptx | ✅ | ✅ | ✅ | PDF | PptxReader / PptxWriter / PptxTemplate / PptxPdfConverter |
| PowerPoint ppt | ✅ | — | — | — | PptReader |
| PDF | ✅ | ✅ | — | — | PdfReader / PdfWriter / PdfFluentDocument |
| Markdown | ✅ | ✅ | — | HTML、Word、PDF | Markdown 系列 |
| RTF | ✅ | ✅ | ✅ | — | RtfDocument / RtfWriter |
| ODS | ✅ | ✅ | — | — | OdsReader / OdsWriter |
| EML | ✅ | ✅ | — | — | EmlReader / EmlWriter |
| MSG | ✅ | — | — | — | MsgReader |
| iCalendar | ✅ | ✅ | — | — | ICalReader / ICalWriter |
| vCard | ✅ | ✅ | — | — | VCardReader / VCardWriter |
| EPUB | ✅ | ✅ | — | — | EpubReader / EpubWriter |
| XPS | ✅ | ✅ | — | — | XpsReader / XpsWriter |

## 目录结构

```text
Pek.Office/
├── Calendar/
├── Epub/
├── Excel/
├── Mail/
├── Markdown/
├── Ods/
├── OLE2/
├── PDF/
├── PPT/
├── Rtf/
├── VCard/
├── Word/
├── Xps/
└── OfficeFactory.cs
```

## 测试覆盖

仓库已提供独立的 xUnit 测试项目，覆盖以下主要模块：

- ExcelReader / ExcelWriter / ExcelTemplate / BiffReader / BiffWriter
- CfbDocument、DocReader、PptReader
- Eml、Msg、ICal、VCard、Epub
- Markdown、Rtf、Ods、Xps
- WordHtmlConverter、WordPdfConverter、PptxPdfConverter
- OfficeFactory 集成测试

运行测试：

```bash
dotnet test XUnitTest/XUnitTest.csproj
```

## 已知边界

- WordPdfConverter 和 PptxPdfConverter 当前属于内容映射型转换，目标是稳定导出文本和基本结构，不追求 Office 原生渲染级还原
- WordPdfConverter.ConvertToImages 和 PptxPdfConverter.ConvertToImages 当前按设计抛出 NotSupportedException
- OfficeFactory.CreateReader 的返回值类型会随扩展名变化，调用方需要按具体格式进行转换或继续处理

## 适用场景

- 后台报表导出
- 文档批量导入导出
- 旧版 Office 二进制格式解析
- 邮件、日历、联系人、电子书等结构化文档处理
- 轻量级文档格式转换与内容提取