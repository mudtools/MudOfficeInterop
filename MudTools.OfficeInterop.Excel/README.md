# MudTools.OfficeInterop.Excel

Excel 操作模块，提供完整的 Excel 应用程序操作接口。

## 项目概述

MudTools.OfficeInterop.Excel 是专门用于操作 Microsoft Excel 应用程序的 .NET 封装库。该模块提供了完整的 Excel 应用程序操作接口，包括工作簿、工作表、单元格等对象的便捷操作，以及图表、数据透视表等高级功能封装。

通过使用本模块，开发者可以避免直接处理复杂的 Excel COM 交互，从而更专注于业务逻辑的实现。

## 主要功能

- 完整的 Excel 应用程序操作接口
- 工作簿、工作表、单元格等对象的便捷操作
- 图表、数据透视表等高级功能封装
- 格式设置、样式管理等功能

## 支持的框架

- .NET Framework 4.6.2
- .NET Framework 4.7
- .NET Framework 4.8
- .NET Standard 2.1

## 安装

```xml
<PackageReference Include="MudTools.OfficeInterop.Excel" Version="1.1.0" />
```

## 核心组件

### ExcelFactory

[ExcelFactory](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Excel/ExcelFactory.cs#L22-L152) 是用于创建和操作 Excel 应用程序实例的工厂类，提供以下方法：

- `Connection` - 通过现有 COM 对象连接到已运行的 Excel 应用程序实例
- `CreateInstance` - 通过 ProgID 创建特定版本的 Excel 应用程序实例
- `BlankWorkbook` - 创建新的空白 Excel 工作簿
- `CreateFrom` - 基于模板创建新的 Excel 工作簿
- `Open` - 打开现有的 Excel 工作簿文件

## 使用示例

### 基本操作

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

### 从模板创建 Excel 工作簿

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

### 读取 Excel 数据

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

### Excel 图表操作

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

## 许可证

本项目采用双重许可证模式：

- [MIT 许可证](../../LICENSE-MIT)
- [Apache 许可证 2.0](../../LICENSE-APACHE)

## 免责声明

本项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。

不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任。