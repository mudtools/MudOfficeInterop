# MudTools.OfficeInterop.Word

Word 操作模块，提供完整的 Word 文档操作接口。

## 项目概述

MudTools.OfficeInterop.Word 是专门用于操作 Microsoft Word 应用程序的 .NET 封装库。该模块提供了完整的 Word 文档操作接口，包括文档内容、样式、格式等管理功能，以及表格、图片等元素的操作封装。

通过使用本模块，开发者可以避免直接处理复杂的 Word COM 交互，从而更专注于业务逻辑的实现。

## 主要功能

- Word 文档操作接口
- 文档内容、样式、格式等管理功能
- 表格、图片等元素的操作封装

## 支持的框架

- .NET Framework 4.6.2
- .NET Framework 4.7
- .NET Framework 4.8
- .NET Standard 2.1
- .NET 6.0-windows
- .NET 7.0-windows
- .NET 8.0-windows
- .NET 9.0-windows

## 安装

```xml
<PackageReference Include="MudTools.OfficeInterop.Word" Version="1.1.8" />
```

## 核心组件

### WordFactory

[WordFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/WordFactory.cs#L15-L97) 是用于创建和操作 Word 应用程序实例的工厂类，提供以下方法：

- `Connection` - 通过现有 COM 对象连接到已运行的 Word 应用程序实例
- `BlankWorkbook` - 创建新的空白 Word 文档
- `CreateFrom` - 基于模板创建新的 Word 文档
- `Open` - 打开现有的 Word 文档文件

## 使用示例

### 基本操作

```csharp
// 创建 Word 应用程序实例
using var app = WordFactory.CreateApplication();
app.Visible = true;

// 创建新文档
var document = app.Documents.Add();

// 添加内容
var range = document.Range();
range.Text = "Hello World!";

// 保存文档
document.SaveAs2(@"C:\temp\example.docx");
```

### 使用模板创建 Word 文档

```csharp
// 基于模板创建文档
using var app = WordFactory.CreateFrom(@"C:\templates\ReportTemplate.dotx");
var document = app.ActiveDocument;

// 替换模板中的占位符
var selection = app.Selection;
selection.Find.Text = "{REPORT_TITLE}";
selection.Find.Replacement.Text = "季度销售报告";
selection.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);

// 添加表格
var table = document.Tables.Add(document.Range(document.Content.End - 1, document.Content.End - 1), 3, 3);
table.Cell(1, 1).Range.Text = "产品";
table.Cell(1, 2).Range.Text = "销量";
table.Cell(1, 3).Range.Text = "收入";

app.Quit();
```

### Word 文档格式化

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加标题
var titleRange = document.Range();
titleRange.Text = "文档标题\n";
titleRange.Font.Bold = 1;
titleRange.Font.Size = 16;
titleRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

// 添加段落
var paraRange = document.Range(document.Content.End - 1, document.Content.End - 1);
paraRange.Text = "这是文档的内容段落，包含一些示例文本。\n";
paraRange.Font.Bold = 0;
paraRange.Font.Size = 12;

// 添加列表
var listRange = document.Range(document.Content.End - 1, document.Content.End - 1);
listRange.Text = "项目1\n项目2\n项目3\n";
listRange.ListFormat.ApplyBulletDefault();

document.SaveAs2(@"C:\documents\FormattedDocument.docx");
app.Quit();
```

## 许可证

本项目采用双重许可证模式：

- [MIT 许可证](../../LICENSE-MIT)

## 免责声明

本项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。

不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任。