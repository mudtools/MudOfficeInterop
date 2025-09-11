# MudTools.OfficeInterop

一个针对 Microsoft Office 应用程序的 .NET 封装库，旨在简化 Office COM 组件的使用。

该库为开发者提供了一套现代化、面向对象的 API，用于操作 Microsoft Office 应用程序（Excel、Word、PowerPoint）。通过使用本库，开发者可以避免直接处理复杂的 COM 交互，从而更专注于业务逻辑的实现。

## 项目概述

MudTools.OfficeInterop 是一套针对 Microsoft Office 应用程序（包括 Excel、Word、PowerPoint 和 VBE）的 .NET 封装库。该项目通过提供简洁、统一的 API 接口，降低了直接使用 Office COM 组件的复杂性，使开发者能够更轻松地在 .NET 应用程序中集成和操作 Office 文档。

### 项目目标

本项目的主要目标是：

1. **简化 Office 自动化**：通过封装复杂的 COM 接口，提供更简洁、更易用的 .NET API
2. **提高开发效率**：减少开发者在 Office 自动化方面所需的时间和精力
3. **增强代码可维护性**：通过面向对象的设计和清晰的接口，使代码更易于理解和维护
4. **提供完整功能覆盖**：支持 Office 应用程序的常用功能，包括文档创建、编辑、格式化等
5. **确保类型安全**：利用 .NET 的类型系统，减少运行时错误

### 适用场景

MudTools.OfficeInterop 适用于以下场景：

- 企业报表生成和数据处理
- 批量文档处理和格式化
- Office 插件开发
- 自动化办公应用
- 数据导入/导出功能
- 文档模板处理

### 设计理念

本项目遵循以下设计理念：

1. **简洁性**：提供直观、易用的 API，降低学习成本
2. **一致性**：在不同 Office 应用程序间保持相似的接口设计
3. **可扩展性**：允许开发者在需要时访问底层 COM 对象
4. **资源管理**：通过实现 IDisposable 接口，确保正确释放 COM 资源
5. **兼容性**：支持多个 .NET Framework 版本和不同版本的 Office

## 功能模块

### 核心模块 (MudTools.OfficeInterop)
- 提供 Office 应用程序的基础接口和通用功能
- 封装 Office 核心组件的常用操作
- 为其他 Office 应用程序模块提供基础支撑
- 提供 Office UI 相关组件的封装，包括功能区(Ribbon)和自定义任务窗格(CTP)

### Excel 模块 (MudTools.OfficeInterop.Excel)
- 完整的 Excel 应用程序操作接口
- 工作簿、工作表、单元格等对象的便捷操作
- 图表、数据透视表等高级功能封装
- 格式设置、样式管理等功能

### Word 模块 (MudTools.OfficeInterop.Word)
- Word 文档操作接口
- 文档内容、样式、格式等管理功能
- 表格、图片等元素的操作封装

### PowerPoint 模块 (MudTools.OfficeInterop.PowerPoint)
- PowerPoint 演示文稿操作接口
- 幻灯片、母版、动画等对象的管理
- 演示文稿的创建、编辑和格式化功能

### VBE 模块 (MudTools.OfficeInterop.Vbe)
- Visual Basic Editor 相关功能封装
- 宏、代码模块、项目等对象的操作接口

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

该项目依赖于 Microsoft Office COM 组件，使用前需要确保系统中已安装相应版本的 Microsoft Office。

```xml
<PackageReference Include="MudTools.OfficeInterop" Version="1.1.2" />
<PackageReference Include="MudTools.OfficeInterop.Excel" Version="1.1.2" />
```

| 模块 | 当前版本 | 开源协议 | 
|---|---|---|
| [![OfficeInterop-Core](https://img.shields.io/badge/OfficeInterop-1.1.2-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/badge/%E7%89%88%E6%9C%AC-1.1.2-blue&label=Version&logo=nuget)](https://www.nuget.org/packages/MudTools.OfficeInterop/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE)
| [![OfficeInterop-Excel](https://img.shields.io/badge/OfficeInteropExcel-1.1.2-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/badge/%E7%89%88%E6%9C%AC-1.1.2-blue&label=Version&logo=nuget)](https://www.nuget.org/packages/MudTools.OfficeInterop.Excel/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE)
| [![OfficeInterop-Word](https://img.shields.io/badge/OfficeInteropWord-1.1.2-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/badge/%E7%89%88%E6%9C%AC-1.1.2-blue&label=Version&logo=nuget)](https://www.nuget.org/packages/MudTools.OfficeInterop.Word/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE)
| [![OfficeInterop-PowerPoint](https://img.shields.io/badge/OfficeInteropPowerPoint-1.1.2-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/badge/%E7%89%88%E6%9C%AC-1.1.2-blue&label=Version&logo=nuget)](https://www.nuget.org/packages/MudTools.OfficeInterop.PowerPoint/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE)
| [![OfficeInterop-Vbe](https://img.shields.io/badge/OfficeInteropVbe-1.1.2-success.svg)](https://gitee.com/mudtools/OfficeInterop) | [![Nuget](https://img.shields.io/badge/%E7%89%88%E6%9C%AC-1.1.2-blue&label=Version&logo=nuget)](https://www.nuget.org/packages/MudTools.OfficeInterop.Vbe/) | [![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://gitee.com/mudtools/OfficeInterop/blob/master/LICENSE)

## 工厂类使用说明

本项目提供多个工厂类用于创建和操作 Office 应用程序对象：

- [OfficeUIFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop/OfficeUIFactory.cs#L16-L51) - 用于创建 Office UI 相关组件，如功能区(Ribbon)和自定义任务窗格(CTP)
- [ExcelFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Excel/ExcelFactory.cs#L22-L152) - 用于创建和操作 Excel 应用程序实例
- [WordFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/WordFactory.cs#L15-L97) - 用于创建和操作 Word 应用程序实例
- [PowerPointFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.PowerPoint/PowerPointFactory.cs#L15-L74) - 用于创建和操作 PowerPoint 应用程序实例

所有工厂类都提供多种创建应用程序实例的方法：
- `Connection` - 通过现有 COM 对象连接到已运行的应用程序实例
- `BlankWorkbook` - 创建新的空白文档/工作簿/演示文稿
- `CreateFrom` - 基于模板创建新的文档/工作簿/演示文稿
- `Open` - 打开现有的文档/工作簿/演示文稿
- `CreateInstance` - (仅 ExcelFactory) 通过 ProgID 创建特定版本的应用程序实例

## 使用示例

### Excel 操作示例

#### 基本操作

```csharp
// 创建 Excel 应用程序实例
using var app = ExcelFactory.CreateApplication();
app.Visible = true;

// 添加工作簿
var workbook = app.Workbooks.Add();
var worksheet = workbook.Worksheets.Add();

// 操作单元格
worksheet.Range["A1"].Value = "Hello";
worksheet.Range["B1"].Value = "World";

// 保存工作簿
workbook.SaveAs(@"C:\temp\example.xlsx");
```

#### 从模板创建 Excel 工作簿

```csharp
// 基于模板创建工作簿
using var app = ExcelFactory.CreateFrom(@"C:\templates\ReportTemplate.xltx");
var worksheet = app.GetActiveSheet();

// 填充数据
worksheet.Range["A1"].Value = "销售报告";
worksheet.Range["A2"].Value = DateTime.Now.ToString("yyyy-MM-dd");

// 保存并关闭
app.ActiveWorkbook.SaveAs(@"C:\reports\SalesReport.xlsx");
app.Quit();
```

#### 读取 Excel 数据

```csharp
// 打开现有工作簿
using var app = ExcelFactory.Open(@"C:\data\SalesData.xlsx");
var worksheet = app.Worksheets[1];

// 读取数据范围
var dataRange = worksheet.Range["A1:D100"];
var rowCount = dataRange.Rows.Count;
var columnCount = dataRange.Columns.Count;

// 处理数据
for (int row = 1; row <= rowCount; row++)
{
    for (int col = 1; col <= columnCount; col++)
    {
        var cellValue = dataRange.Cells[row, col].Value?.ToString();
        Console.WriteLine($"Row {row}, Column {col}: {cellValue}");
    }
}

app.Quit();
```

#### Excel 图表操作

```csharp
using var app = ExcelFactory.BlankWorkbook();
var worksheet = app.GetActiveSheet();

// 添加示例数据
worksheet.Range["A1"].Value = "月份";
worksheet.Range["B1"].Value = "销售额";
worksheet.Range["A2"].Value = "一月";
worksheet.Range["B2"].Value = 10000;
worksheet.Range["A3"].Value = "二月";
worksheet.Range["B3"].Value = 15000;
worksheet.Range["A4"].Value = "三月";
worksheet.Range["B4"].Value = 12000;

// 创建图表
var chartObjects = worksheet.ChartObjects();
var chartObject = chartObjects.Add(100, 50, 300, 200);
var chart = chartObject.Chart;

// 设置图表数据源
chart.SetSourceData(worksheet.Range["A1:B4"]);
chart.ChartType = XlChartType.xlColumnClustered;

app.ActiveWorkbook.SaveAs(@"C:\charts\SalesChart.xlsx");
app.Quit();
```

### Word 操作示例

#### 基本操作

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

#### 使用模板创建 Word 文档

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

#### Word 文档格式化

```csharp
using var app = WordFactory.BlankWorkbook();
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

### PowerPoint 操作示例

#### 创建演示文稿

```csharp
// 创建 PowerPoint 应用程序实例
using var app = PowerPointFactory.CreateApplication();
app.Visible = true;

// 创建新演示文稿
var presentation = app.Presentations.Add();

// 添加幻灯片
var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);

// 设置标题
slide.Shapes.Title.TextFrame.TextRange.Text = "欢迎使用 PowerPoint";

// 添加内容
slide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "这是演示文稿的内容部分";

// 保存演示文稿
presentation.SaveAs(@"C:\presentations\example.pptx");
```

#### 操作现有演示文稿

``csharp
// 打开现有演示文稿
using var app = PowerPointFactory.Open(@"C:\presentations\ExistingPresentation.pptx");
var presentation = app.ActivePresentation;

// 遍历所有幻灯片
foreach (var slide in presentation.Slides)
{
    Console.WriteLine($"幻灯片 {slide.SlideIndex}: {slide.Name}");
    
    // 修改幻灯片内容
    if (slide.Shapes.HasTitle == MsoTriState.msoTrue)
    {
        slide.Shapes.Title.TextFrame.TextRange.Text += " - 已更新";
    }
}

// 添加新幻灯片
var newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, 
                                      PowerPoint.PpSlideLayout.ppLayoutText);

newSlide.Shapes.Title.TextFrame.TextRange.Text = "新幻灯片";
newSlide.Shapes.Placeholders[2].TextFrame.TextRange.Text = "这是新增的幻灯片内容";

presentation.Save();
app.Quit();
```

#### PowerPoint 格式化和动画

```csharp
using var app = PowerPointFactory.BlankWorkbook();
var presentation = app.ActivePresentation;

// 添加幻灯片
var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

// 添加形状
var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 100, 100, 200, 100);
shape.TextFrame.TextRange.Text = "示例形状";

// 设置形状格式
shape.Fill.ForeColor.RGB = 0x00FF00; // 绿色填充
shape.Line.ForeColor.RGB = 0xFF0000; // 红色边框

// 添加动画
var animation = shape.AnimationSettings;
animation.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFade;
animation.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnClick;

presentation.SaveAs(@"C:\presentations\AnimatedPresentation.pptx");
app.Quit();
```

### Office UI 操作示例

#### 使用自定义任务窗格

```csharp
// 创建自定义任务窗格
var ctpFactory = OfficeUIFactory.CreateCTPFactory(officeCTPFactory);
var ctp = ctpFactory.CreateCTP("MyAddin.UserControl", "我的任务窗格");

// 设置任务窗格属性
ctp.Visible = true;
ctp.Width = 200;

// 显示任务窗格
ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
```

#### 使用功能区控件

```csharp
// 处理功能区控件事件
public void OnRibbonButtonClicked(IRibbonControl control)
{
    switch (control.Id)
    {
        case "buttonNewDocument":
            // 创建新文档
            using var app = ExcelFactory.BlankWorkbook();
            break;
        case "buttonOpenDocument":
            // 打开文档
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                using var app = ExcelFactory.Open(openFileDialog.FileName);
            }
            break;
    }
}
```

## 许可证

本项目采用MIT许可证模式：

- [MIT 许可证](LICENSE-MIT)

## 免责声明

本项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。

不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任。