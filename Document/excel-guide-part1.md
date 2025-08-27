# Excel 操作指南（第一部分）：核心应用与窗口管理

## 适用场景与解决问题

本指南适用于需要通过 .NET 程序操作 Excel 应用程序的开发者，解决以下问题：
- 如何启动和连接 Excel 应用程序
- 如何管理 Excel 工作簿和窗口
- 如何简化 Excel 自动化操作
- 如何避免 COM 对象管理的复杂性

## ExcelFactory - Excel 应用程序入口点

[ExcelFactory](file:///D:/Repos/OfficeInterop/MudTools.OfficeInterop.Excel/ExcelFactory.cs#L22-L152) 是创建和操作 Excel 应用程序的静态工厂类，提供了多种创建 Excel 实例的方法。

### 主要方法

#### 1. BlankWorkbook() - 创建空白工作簿
```csharp
// 创建新的空白工作簿
var excelApp = ExcelFactory.BlankWorkbook();
// 现在可以对工作簿进行操作
excelApp.GetActiveSheet().Cells[1, 1].Value = "Hello World";
```

#### 2. CreateFrom(string templatePath) - 基于模板创建工作簿
```csharp
// 基于模板创建工作簿
var excelApp = ExcelFactory.CreateFrom(@"C:\Templates\ReportTemplate.xltx");
// 新工作簿将继承模板的格式、样式、公式等
```

#### 3. Open(string filePath) - 打开现有工作簿
```csharp
// 打开现有工作簿
var excelApp = ExcelFactory.Open(@"C:\Data\SalesReport.xlsx");
// 现在可以读取和修改现有数据
var value = excelApp.GetActiveSheet().Cells[1, 1].Value;
```

#### 4. Connection(object comObj) - 连接现有 Excel 实例
```csharp
// 连接到现有的 Excel 应用程序实例
var excelApp = ExcelFactory.Connection(comObject);
```

#### 5. CreateInstance(string progId) - 创建特定版本实例
```csharp
// 根据 ProgID 创建 Excel 应用程序的新实例
var excelApp = ExcelFactory.CreateInstance("Excel.Application.16");
```

## IExcelApplication - Excel 应用程序核心接口

[IExcelApplication](file:///D:/Repos/OfficeInterop/MudTools.OfficeInterop.Excel/Core/IExcelApplication.cs#L12-L1129) 是操作 Excel 应用程序的核心接口，提供了对 Excel 应用程序的全面控制。

### 基础属性管理

```csharp
// 设置应用程序属性
excelApp.DisplayAlerts = false; // 禁用警告对话框
excelApp.ScreenUpdating = false; // 禁用屏幕更新以提高性能
excelApp.Calculation = XlCalculation.xlCalculationManual; // 手动计算模式

// 获取系统信息
string version = excelApp.Version;
int memoryFree = excelApp.MemoryFree;
```

### 工作簿管理

```csharp
// 获取工作簿集合
var workbooks = excelApp.Workbooks;

// 获取活动工作簿
var activeWorkbook = excelApp.ActiveWorkbook;

// 获取特定工作簿
var workbook = excelApp.GetWorkbook("MyWorkbook.xlsx");
```

### 工作表管理

```csharp
// 获取活动工作表
var activeSheet = excelApp.ActiveSheet;

// 获取工作表集合
var sheets = excelApp.Sheets;

// 获取单元格区域
var range = excelApp.Range("A1:B10");
```

### 计算与公式

```csharp
// 手动计算
excelApp.Calculate();

// 计算特定工作表
excelApp.CalculateWorksheet(worksheet);

// 计算公式
object result = excelApp.Evaluate("=SUM(A1:A10)");
double sum = excelApp.EvaluateToNumber("=SUM(1, 2, 3)");
```

### 用户界面控制

```csharp
// 窗口操作
excelApp.Minimize();
excelApp.Maximize();
excelApp.Restore();

// 显示设置
excelApp.DisplayFullScreen = true;
excelApp.DisplayFormulaBar = false;
```

## IExcelWorkbook - 工作簿操作接口

[IExcelWorkbook](file:///D:/Repos/OfficeInterop/MudTools.OfficeInterop.Excel/Core/IExcelWorkbook.cs#L15-L395) 提供对 Excel 工作簿的全面管理功能。

### 工作簿基础操作

```csharp
// 保存工作簿
workbook.Save();

// 另存为
workbook.SaveAs(@"C:\Output\NewFile.xlsx");

// 关闭工作簿
workbook.Close(saveChanges: true);
```

### 工作表管理

```csharp
// 获取工作表数量
int count = workbook.WorksheetCount;

// 获取特定工作表
var worksheet = workbook.GetWorksheet(1);
var namedSheet = workbook.GetWorksheet("Sheet1");

// 添加新工作表
var newSheet = workbook.AddWorksheet();

// 删除工作表
workbook.DeleteWorksheet(sheetToDelete);
```

### 工作簿保护

```csharp
// 保护工作簿
workbook.Protect("password");

// 取消保护
workbook.Unprotect("password");

// 保护所有工作表
workbook.ProtectAllWorksheets("password");
```

### 高级功能

```csharp
// 导出为PDF
workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, @"C:\Output\Report.pdf");

// 计算所有公式
workbook.CalculateAll();

// 刷新所有数据
workbook.RefreshAll();
```

## IExcelWindow - 窗口管理接口

[IExcelWindow](file:///D:/Repos/OfficeInterop/MudTools.OfficeInterop.Excel/Core/IExcelWindow.cs#L12-L277) 提供对 Excel 窗口的详细控制。

### 窗口属性设置

```csharp
// 窗口状态
window.WindowState = XlWindowState.xlMaximized;

// 显示比例
window.Zoom = 150; // 150%

// 视图类型
window.View = XlWindowView.xlPageBreakPreview;
```

### 窗口显示选项

```csharp
// 网格线显示
window.DisplayGridlines = true;

// 行列标题显示
window.DisplayHeadings = true;

// 公式显示
window.DisplayFormulas = false;

// 零值显示
window.DisplayZeros = true;
```

### 窗口分割与冻结

```csharp
// 冻结窗格
window.FreezePanes = true;

// 拆分窗口
window.Split = true;
window.SplitRow = 5;
window.SplitColumn = 3;
```

### 滚动操作

```csharp
// 大范围滚动
window.LargeScroll(down: 10, right: 5);

// 小范围滚动
window.SmallScroll(down: 1, right: 1);

// 滚动到指定区域
window.ScrollToRange("A100");
```

### 窗口选择与导航

```csharp
// 选择指定范围
window.SelectRange("B2:D10");

// 获取可见区域
var visibleRange = window.VisibleRange;

// 获取选中工作表
var selectedSheets = window.SelectedSheets;
```

## 最佳实践示例

### 完整的工作簿操作示例

```csharp
// 创建新的 Excel 应用程序和工作簿
using var excelApp = ExcelFactory.BlankWorkbook();

try
{
    // 获取活动工作表
    var worksheet = excelApp.ActiveSheet;
    
    // 写入数据
    worksheet.Cells[1, 1].Value = "产品名称";
    worksheet.Cells[1, 2].Value = "销量";
    worksheet.Cells[1, 3].Value = "单价";
    worksheet.Cells[1, 4].Value = "金额";
    
    // 填充数据
    string[] products = { "产品A", "产品B", "产品C" };
    int[] sales = { 100, 200, 150 };
    double[] prices = { 10.5, 15.0, 12.8 };
    
    for (int i = 0; i < products.Length; i++)
    {
        worksheet.Cells[i + 2, 1].Value = products[i];
        worksheet.Cells[i + 2, 2].Value = sales[i];
        worksheet.Cells[i + 2, 3].Value = prices[i];
        worksheet.Cells[i + 2, 4].Formula = $"=B{i + 2}*C{i + 2}";
    }
    
    // 格式化标题行
    var headerRange = worksheet.Range("A1:D1");
    headerRange.Font.Bold = true;
    headerRange.Interior.Color = Color.LightBlue;
    
    // 自动调整列宽
    worksheet.Columns.AutoFit();
    
    // 保存文件
    excelApp.ActiveWorkbook.SaveAs(@"C:\Output\SalesReport.xlsx");
}
finally
{
    // 关闭应用程序
    excelApp.Quit();
}
```

### 窗口管理示例

```csharp
// 窗口设置示例
using var excelApp = ExcelFactory.Open(@"C:\Data\Report.xlsx");

var window = excelApp.ActiveWindow;

// 设置窗口属性
window.WindowState = XlWindowState.xlMaximized;
window.DisplayGridlines = false;
window.DisplayHeadings = true;
window.Zoom = 90;

// 冻结首行
window.SplitRow = 1;
window.FreezePanes = true;

// 滚动到数据区域
window.ScrollToRange("A100");
```

## 总结

通过使用 ExcelFactory 和相关接口，开发者可以：
1. 简化 Excel 应用程序的创建和管理
2. 避免手动处理 COM 对象生命周期
3. 使用强类型接口提高代码可读性和安全性
4. 更好地控制 Excel 工作簿和窗口行为
5. 提高开发效率和代码维护性

这些接口提供了对 Excel 核心功能的全面封装，使开发者能够专注于业务逻辑而不是底层的 COM 交互细节。