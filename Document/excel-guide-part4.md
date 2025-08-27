# Excel 操作指南（第四部分）：页面设置和打印预览

## 适用场景与解决问题

本指南适用于需要在 Excel 工作表中进行页面设置和打印预览操作的开发者，解决以下问题：
- 如何配置页面方向、纸张大小和缩放比例
- 如何设置页边距和居中方式
- 如何自定义页眉页脚内容
- 如何配置打印选项和区域
- 如何进行打印预览和打印操作
- 如何简化复杂的页面布局和打印任务

## IExcelPageSetup - 页面设置操作接口

[IExcelPageSetup](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Excel/Core/IExcelPageSetup.cs#L15-L367) 用于管理 Excel 工作表的页面设置。

### 页面设置基础操作

```csharp
// 获取或设置页面方向（纵向或横向）
XlPageOrientation orientation = pageSetup.Orientation;
pageSetup.Orientation = XlPageOrientation.xlLandscape;

// 获取或设置纸张大小
XlPaperSize paperSize = pageSetup.PaperSize;
pageSetup.PaperSize = XlPaperSize.xlPaperA4;

// 获取或设置页面缩放比例（10-400）
int zoom = pageSetup.Zoom;
pageSetup.Zoom = 100;

// 获取或设置是否适合页面宽度
int fitToPagesWide = pageSetup.FitToPagesWide;
pageSetup.FitToPagesWide = 1;

// 获取或设置是否适合页面高度
int fitToPagesTall = pageSetup.FitToPagesTall;
pageSetup.FitToPagesTall = 1;
```

### 页边距设置操作

```csharp
// 获取或设置左边距（英寸）
double leftMargin = pageSetup.LeftMargin;
pageSetup.LeftMargin = 0.5;

// 获取或设置右边距（英寸）
double rightMargin = pageSetup.RightMargin;
pageSetup.RightMargin = 0.5;

// 获取或设置上边距（英寸）
double topMargin = pageSetup.TopMargin;
pageSetup.TopMargin = 0.75;

// 获取或设置下边距（英寸）
double bottomMargin = pageSetup.BottomMargin;
pageSetup.BottomMargin = 0.75;

// 获取或设置页眉边距（英寸）
double headerMargin = pageSetup.HeaderMargin;
pageSetup.HeaderMargin = 0.3;

// 获取或设置页脚边距（英寸）
double footerMargin = pageSetup.FooterMargin;
pageSetup.FooterMargin = 0.3;

// 获取或设置居中方式（水平居中）
bool centerHorizontally = pageSetup.CenterHorizontally;
pageSetup.CenterHorizontally = true;

// 获取或设置居中方式（垂直居中）
bool centerVertically = pageSetup.CenterVertically;
pageSetup.CenterVertically = true;
```

### 页眉页脚设置操作

```csharp
// 获取或设置左页眉内容
string leftHeader = pageSetup.LeftHeader;
pageSetup.LeftHeader = "公司名称";

// 获取或设置中页眉内容
string centerHeader = pageSetup.CenterHeader;
pageSetup.CenterHeader = "报告标题";

// 获取或设置右页眉内容
string rightHeader = pageSetup.RightHeader;
pageSetup.RightHeader = "机密";

// 获取或设置左页脚内容
string leftFooter = pageSetup.LeftFooter;
pageSetup.LeftFooter = "作者：张三";

// 获取或设置中页脚内容
string centerFooter = pageSetup.CenterFooter;
pageSetup.CenterFooter = "部门：销售部";

// 获取或设置右页脚内容
string rightFooter = pageSetup.RightFooter;
pageSetup.RightFooter = "第 &P 页，共 &N 页";
```

### 打印选项设置

```csharp
// 获取或设置打印顺序
XlOrder order = pageSetup.Order;
pageSetup.Order = XlOrder.xlDownThenOver;

// 获取或设置是否以草稿模式打印
bool draft = pageSetup.Draft;
pageSetup.Draft = false;

// 获取或设置是否打印网格线
bool printGridlines = pageSetup.PrintGridlines;
pageSetup.PrintGridlines = true;

// 获取或设置是否打印行列标题
bool printHeadings = pageSetup.PrintHeadings;
pageSetup.PrintHeadings = true;

// 获取或设置打印标题行
string printTitleRows = pageSetup.PrintTitleRows;
pageSetup.PrintTitleRows = "$1:$1";

// 获取或设置打印标题列
string printTitleColumns = pageSetup.PrintTitleColumns;
pageSetup.PrintTitleColumns = "$A:$A";

// 获取或设置打印区域
string printArea = pageSetup.PrintArea;
pageSetup.PrintArea = "$A$1:$E$20";

// 获取或设置是否从第一页开始编号
int firstPageNumber = pageSetup.FirstPageNumber;
pageSetup.FirstPageNumber = 1;
```

### 页面设置操作方法

```csharp
// 应用页面设置
pageSetup.Apply();

// 重置页面设置为默认值
pageSetup.Reset();

// 复制页面设置
pageSetup.Copy(sourcePageSetup);

// 设置自定义页眉页脚
pageSetup.SetCustomHeaderFooter(1, 1, "自定义页眉内容");
```

## IExcelPrintPreview - 打印预览操作接口

[IExcelPrintPreview](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Excel/Core/IExcelPrintPreview.cs#L13-L204) 用于管理 Excel 工作表的打印预览功能。

### 打印预览基础操作

```csharp
// 获取或设置打印预览的缩放比例（10-400）
int zoom = printPreview.Zoom;
printPreview.Zoom = 100;

// 获取或设置是否显示页眉
bool showHeaders = printPreview.ShowHeaders;
printPreview.ShowHeaders = true;

// 获取或设置是否显示页脚
bool showFooters = printPreview.ShowFooters;
printPreview.ShowFooters = true;

// 获取或设置是否显示网格线
bool showGridlines = printPreview.ShowGridlines;
printPreview.ShowGridlines = true;

// 获取或设置是否显示行列标题
bool showHeadings = printPreview.ShowHeadings;
printPreview.ShowHeadings = true;
```

### 打印预览页面设置

```csharp
// 获取或设置页面方向（纵向或横向）
int orientation = printPreview.Orientation;
printPreview.Orientation = 2; // 横向

// 获取或设置纸张大小
int paperSize = printPreview.PaperSize;
printPreview.PaperSize = 9; // A4

// 获取或设置页边距（英寸）
double leftMargin = printPreview.LeftMargin;
printPreview.LeftMargin = 0.5;

double rightMargin = printPreview.RightMargin;
printPreview.RightMargin = 0.5;

double topMargin = printPreview.TopMargin;
printPreview.TopMargin = 0.75;

double bottomMargin = printPreview.BottomMargin;
printPreview.BottomMargin = 0.75;
```

### 打印预览页眉页脚设置

```csharp
// 获取或设置页眉内容
string leftHeader = printPreview.LeftHeader;
printPreview.LeftHeader = "公司名称";

string centerHeader = printPreview.CenterHeader;
printPreview.CenterHeader = "报告标题";

string rightHeader = printPreview.RightHeader;
printPreview.RightHeader = "机密";

// 获取或设置页脚内容
string leftFooter = printPreview.LeftFooter;
printPreview.LeftFooter = "作者：张三";

string centerFooter = printPreview.CenterFooter;
printPreview.CenterFooter = "部门：销售部";

string rightFooter = printPreview.RightFooter;
printPreview.RightFooter = "第 &P 页，共 &N 页";
```

### 打印预览操作方法

```csharp
// 显示打印预览窗口
printPreview.Show();

// 刷新打印预览显示
printPreview.Refresh();

// 打印当前预览的内容
printPreview.Print(2, true); // 打印2份，逐份打印

// 导出预览为PDF文件
printPreview.ExportToPDF(@"C:\Reports\Report.pdf");
```

## 实际应用示例

### 配置工作表页面设置

```csharp
// 配置工作表页面设置
using var excelApp = ExcelFactory.BlankWorkbook();
var workbook = excelApp.ActiveWorkbook;
var worksheet = excelApp.GetActiveSheet();

try
{
    // 填充一些示例数据
    worksheet.Cells[1, 1].Value = "产品销售报告";
    worksheet.Cells[2, 1].Value = "产品名称";
    worksheet.Cells[2, 2].Value = "销售数量";
    worksheet.Cells[2, 3].Value = "销售额";
    
    for (int i = 3; i <= 20; i++)
    {
        worksheet.Cells[i, 1].Value = $"产品{i-2}";
        worksheet.Cells[i, 2].Value = new Random().Next(100, 1000);
        worksheet.Cells[i, 3].Value = new Random().Next(1000, 10000);
    }
    
    // 获取页面设置对象
    var pageSetup = worksheet.PageSetup;
    
    // 设置页面方向为横向
    pageSetup.Orientation = XlPageOrientation.xlLandscape;
    
    // 设置纸张大小为A4
    pageSetup.PaperSize = XlPaperSize.xlPaperA4;
    
    // 设置缩放为100%
    pageSetup.Zoom = 100;
    
    // 设置页边距
    pageSetup.LeftMargin = 0.5;
    pageSetup.RightMargin = 0.5;
    pageSetup.TopMargin = 0.75;
    pageSetup.BottomMargin = 0.75;
    
    // 设置居中方式
    pageSetup.CenterHorizontally = true;
    pageSetup.CenterVertically = false;
    
    // 设置页眉页脚
    pageSetup.LeftHeader = "公司名称";
    pageSetup.CenterHeader = "产品销售报告";
    pageSetup.RightHeader = DateTime.Now.ToString("yyyy-MM-dd");
    pageSetup.CenterFooter = "第 &P 页，共 &N 页";
    
    // 设置打印选项
    pageSetup.PrintGridlines = true;
    pageSetup.PrintHeadings = true;
    pageSetup.PrintTitleRows = "$2:$2"; // 打印标题行
    
    // 应用设置
    pageSetup.Apply();
    
    // 保存文件
    workbook.SaveAs(@"C:\Reports\SalesReport.xlsx");
}
finally
{
    excelApp.Quit();
}
```

### 使用打印预览功能

```csharp
// 使用打印预览功能
using var excelApp = ExcelFactory.Open(@"C:\Reports\SalesReport.xlsx");
var workbook = excelApp.ActiveWorkbook;
var worksheet = excelApp.GetActiveSheet();

try
{
    // 获取打印预览对象
    var printPreview = worksheet.PrintPreview;
    
    // 配置打印预览设置
    printPreview.Zoom = 80;
    printPreview.ShowGridlines = true;
    printPreview.ShowHeadings = true;
    printPreview.LeftHeader = "公司名称";
    printPreview.CenterHeader = "产品销售报告";
    printPreview.RightHeader = DateTime.Now.ToString("yyyy-MM-dd");
    printPreview.CenterFooter = "第 &P 页，共 &N 页";
    
    // 显示打印预览
    printPreview.Show();
    
    // 用户在预览窗口中可以查看效果，确认无误后可以打印
    // printPreview.Print(1, true);
}
finally
{
    excelApp.Quit();
}
```

### 批量处理工作表页面设置

```csharp
// 批量处理工作表页面设置
using var excelApp = ExcelFactory.Open(@"C:\Reports\AnnualReport.xlsx");
var workbook = excelApp.ActiveWorkbook;

try
{
    // 遍历所有工作表并设置统一的页面配置
    foreach (var worksheet in workbook.Worksheets)
    {
        var pageSetup = worksheet.PageSetup;
        
        // 统一设置页面方向
        pageSetup.Orientation = XlPageOrientation.xlPortrait;
        
        // 统一设置纸张大小
        pageSetup.PaperSize = XlPaperSize.xlPaperA4;
        
        // 统一设置页边距
        pageSetup.LeftMargin = 0.75;
        pageSetup.RightMargin = 0.75;
        pageSetup.TopMargin = 1.0;
        pageSetup.BottomMargin = 1.0;
        
        // 统一设置页眉页脚
        pageSetup.CenterHeader = workbook.Name;
        pageSetup.CenterFooter = "第 &P 页";
        
        // 统一设置打印选项
        pageSetup.PrintGridlines = false;
        pageSetup.CenterHorizontally = true;
        
        // 应用设置
        pageSetup.Apply();
    }
    
    // 保存文件
    workbook.Save();
}
finally
{
    excelApp.Quit();
}
```

## 性能优化建议

### 批量页面设置操作

```csharp
// 在操作大量页面设置时禁用屏幕更新
excelApp.ScreenUpdating = false;

try
{
    // 批量操作页面设置
    foreach (var worksheet in workbook.Worksheets)
    {
        var pageSetup = worksheet.PageSetup;
        pageSetup.Orientation = XlPageOrientation.xlLandscape;
        pageSetup.PaperSize = XlPaperSize.xlPaperA4;
        pageSetup.Zoom = 100;
        pageSetup.Apply();
    }
}
finally
{
    excelApp.ScreenUpdating = true;
}
```

### 打印预览性能优化

```csharp
// 在执行打印预览操作时优化性能
excelApp.ScreenUpdating = false;

try
{
    // 执行打印预览相关操作
    var printPreview = worksheet.PrintPreview;
    printPreview.Refresh();
}
finally
{
    excelApp.ScreenUpdating = true;
}
```

## 最佳实践

### 错误处理

```csharp
try
{
    // 操作页面设置
    var pageSetup = worksheet.PageSetup;
    pageSetup.Orientation = XlPageOrientation.xlLandscape;
    pageSetup.Apply();
}
catch (Exception ex)
{
    // 处理异常
    Console.WriteLine($"页面设置操作失败: {ex.Message}");
}
```

### 资源管理

```csharp
// 使用 using 语句确保资源正确释放
using var excelApp = ExcelFactory.BlankWorkbook();
try
{
    var workbook = excelApp.ActiveWorkbook;
    var worksheet = excelApp.GetActiveSheet();
    
    // 执行页面设置和打印预览操作
    ProcessPageSetupAndPreview(worksheet);
    
    // 保存工作簿
    workbook.SaveAs(@"C:\Reports\ProcessedReport.xlsx");
}
finally
{
    excelApp.Quit();
}
```

## 总结

通过使用 IExcelPageSetup 和 IExcelPrintPreview 接口，开发者可以：

1. 精确控制 Excel 工作表的页面布局和打印设置
2. 自定义页眉页脚内容，包括插入页码、日期等动态信息
3. 配置打印选项，如网格线、标题行、打印区域等
4. 预览打印效果，确保输出符合预期
5. 简化复杂的页面设置和打印任务自动化流程

这些接口提供了对 Excel 页面设置和打印预览功能的全面封装，使开发者能够专注于业务逻辑而不是底层的 COM 交互细节。