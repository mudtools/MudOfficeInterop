# 第八篇：页面布局与打印设置详解

## 引言：Excel自动化的"印刷大师"

在Excel自动化开发中，如果说数据是报表的"内容"，格式是报表的"外观"，那么页面布局和打印设置就是报表的"最终呈现"！一个精心设计的报表如果打印效果不佳，就像一幅名画被随意装裱——价值大打折扣。

想象一下这样的场景：你花费了大量心血创建了一份精美的销售报表，数据准确、格式美观、图表生动。但是当需要打印出来呈现在会议上时，却发现表格被截断、页眉页脚错位、页码混乱。这不仅影响了报表的专业性，更可能误导决策者的判断。

MudTools.OfficeInterop.Excel项目就像是专业的"印刷大师"，它通过`IExcelPageSetup`接口提供了完整的页面布局和打印控制功能。从页面方向到纸张大小，从页边距到页眉页脚，每一个打印细节都能得到精确的控制。

本篇将带你探索页面布局和打印设置的奥秘，学习如何通过代码创建专业级的打印文档。准备好让你的Excel自动化报表从"屏幕完美"升级到"打印完美"了吗？

## 页面布局基础

### 页面方向与纸张大小

页面方向（纵向/横向）和纸张大小是页面布局的基础设置。MudTools提供了丰富的枚举类型来支持各种标准纸张规格。

```csharp
public class PageLayoutManager
{
    private readonly IExcelPageSetup _pageSetup;
    
    public PageLayoutManager(IExcelPageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
    }
    
    /// <summary>
    /// 设置标准A4纵向布局
    /// </summary>
    public void SetA4Portrait()
    {
        _pageSetup.Orientation = XlPageOrientation.xlPortrait;
        _pageSetup.PaperSize = XlPaperSize.xlPaperA4;
        _pageSetup.Zoom = 100; // 100%缩放
    }
    
    /// <summary>
    /// 设置A4横向布局，适合宽表格
    /// </summary>
    public void SetA4Landscape()
    {
        _pageSetup.Orientation = XlPageOrientation.xlLandscape;
        _pageSetup.PaperSize = XlPaperSize.xlPaperA4;
        _pageSetup.Zoom = 100;
    }
    
    /// <summary>
    /// 设置自定义纸张大小
    /// </summary>
    public void SetCustomPaperSize(XlPaperSize paperSize, XlPageOrientation orientation)
    {
        _pageSetup.PaperSize = paperSize;
        _pageSetup.Orientation = orientation;
    }
    
    /// <summary>
    /// 设置页面缩放以适应内容
    /// </summary>
    public void SetFitToPage(int pagesWide = 1, int pagesTall = 1)
    {
        _pageSetup.FitToPagesWide = pagesWide;
        _pageSetup.FitToPagesTall = pagesTall;
        _pageSetup.Zoom = false; // 禁用缩放，使用适合页面
    }
    
    /// <summary>
    /// 设置自定义缩放比例
    /// </summary>
    public void SetCustomZoom(int zoomPercentage)
    {
        if (zoomPercentage < 10 || zoomPercentage > 400)
            throw new ArgumentException("缩放比例必须在10-400之间");
            
        _pageSetup.Zoom = zoomPercentage;
        _pageSetup.FitToPagesWide = 1; // 重置适合页面设置
        _pageSetup.FitToPagesTall = 1;
    }
}
```

### 应用场景：财务报表布局设置

```csharp
public class FinancialReportLayout
{
    private readonly PageLayoutManager _layoutManager;
    
    public FinancialReportLayout(IExcelPageSetup pageSetup)
    {
        _layoutManager = new PageLayoutManager(pageSetup);
    }
    
    /// <summary>
    /// 设置资产负债表布局
    /// </summary>
    public void SetBalanceSheetLayout()
    {
        // 资产负债表通常需要横向布局以容纳多列数据
        _layoutManager.SetA4Landscape();
        _layoutManager.SetFitToPage(1, 0); // 适合页面宽度，高度不限
    }
    
    /// <summary>
    /// 设置利润表布局
    /// </summary>
    public void SetIncomeStatementLayout()
    {
        // 利润表可以使用纵向布局
        _layoutManager.SetA4Portrait();
        _layoutManager.SetCustomZoom(90); // 90%缩放以获得更好的可读性
    }
    
    /// <summary>
    /// 设置现金流量表布局
    /// </summary>
    public void SetCashFlowLayout()
    {
        // 现金流量表需要横向布局
        _layoutManager.SetA4Landscape();
        _layoutManager.SetFitToPage(1, 1); // 适合单页
    }
}
```

## 页边距设置

### 标准页边距配置

页边距设置直接影响文档的打印效果和美观度。MudTools提供了精确的页边距控制功能。

```csharp
public class MarginManager
{
    private readonly IExcelPageSetup _pageSetup;
    
    public MarginManager(IExcelPageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
    }
    
    /// <summary>
    /// 设置标准页边距（1英寸）
    /// </summary>
    public void SetStandardMargins()
    {
        _pageSetup.LeftMargin = 1.0;   // 1英寸
        _pageSetup.RightMargin = 1.0;
        _pageSetup.TopMargin = 1.0;
        _pageSetup.BottomMargin = 1.0;
        _pageSetup.HeaderMargin = 0.5; // 页眉边距
        _pageSetup.FooterMargin = 0.5; // 页脚边距
    }
    
    /// <summary>
    /// 设置窄页边距（0.5英寸）
    /// </summary>
    public void SetNarrowMargins()
    {
        _pageSetup.LeftMargin = 0.5;
        _pageSetup.RightMargin = 0.5;
        _pageSetup.TopMargin = 0.5;
        _pageSetup.BottomMargin = 0.5;
        _pageSetup.HeaderMargin = 0.25;
        _pageSetup.FooterMargin = 0.25;
    }
    
    /// <summary>
    /// 设置宽页边距（2英寸）
    /// </summary>
    public void SetWideMargins()
    {
        _pageSetup.LeftMargin = 2.0;
        _pageSetup.RightMargin = 2.0;
        _pageSetup.TopMargin = 2.0;
        _pageSetup.BottomMargin = 2.0;
        _pageSetup.HeaderMargin = 1.0;
        _pageSetup.FooterMargin = 1.0;
    }
    
    /// <summary>
    /// 设置自定义页边距
    /// </summary>
    public void SetCustomMargins(double left, double right, double top, double bottom, 
                                double header = 0.5, double footer = 0.5)
    {
        _pageSetup.LeftMargin = left;
        _pageSetup.RightMargin = right;
        _pageSetup.TopMargin = top;
        _pageSetup.BottomMargin = bottom;
        _pageSetup.HeaderMargin = header;
        _pageSetup.FooterMargin = footer;
    }
    
    /// <summary>
    /// 设置页面居中方式
    /// </summary>
    public void SetCenterAlignment(bool horizontal = true, bool vertical = true)
    {
        _pageSetup.CenterHorizontally = horizontal;
        _pageSetup.CenterVertically = vertical;
    }
}
```

### 应用场景：专业报告页边距设置

```csharp
public class ProfessionalReportMargins
{
    private readonly MarginManager _marginManager;
    
    public ProfessionalReportMargins(IExcelPageSetup pageSetup)
    {
        _marginManager = new MarginManager(pageSetup);
    }
    
    /// <summary>
    /// 设置商业报告页边距
    /// </summary>
    public void SetBusinessReportMargins()
    {
        // 商业报告通常需要较宽的左边距用于装订
        _marginManager.SetCustomMargins(
            left: 1.5,    // 左边距1.5英寸用于装订
            right: 1.0,   // 右边距1英寸
            top: 1.0,     // 上边距1英寸
            bottom: 1.0   // 下边距1英寸
        );
        _marginManager.SetCenterAlignment(horizontal: true, vertical: false);
    }
    
    /// <summary>
    /// 设置学术论文页边距
    /// </summary>
    public void SetAcademicPaperMargins()
    {
        // 学术论文需要标准页边距
        _marginManager.SetStandardMargins();
        _marginManager.SetCenterAlignment(horizontal: true, vertical: true);
    }
    
    /// <summary>
    /// 设置演示文稿页边距
    /// </summary>
    public void SetPresentationMargins()
    {
        // 演示文稿可以使用窄页边距以最大化内容区域
        _marginManager.SetNarrowMargins();
        _marginManager.SetCenterAlignment(horizontal: true, vertical: true);
    }
}
```

## 页眉页脚设计

### 标准页眉页脚设置

页眉页脚是专业文档的重要组成部分，MudTools提供了丰富的页眉页脚设置功能。

```csharp
public class HeaderFooterManager
{
    private readonly IExcelPageSetup _pageSetup;
    
    public HeaderFooterManager(IExcelPageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
    }
    
    /// <summary>
    /// 设置标准页眉：左侧文件名，中间标题，右侧页码
    /// </summary>
    public void SetStandardHeader(string title = "")
    {
        _pageSetup.LeftHeader = "&F";        // 文件名
        _pageSetup.CenterHeader = title;     // 自定义标题
        _pageSetup.RightHeader = "第&P页";   // 页码
    }
    
    /// <summary>
    /// 设置标准页脚：左侧日期，中间公司名，右侧总页数
    /// </summary>
    public void SetStandardFooter(string companyName = "")
    {
        _pageSetup.LeftFooter = "&D";         // 日期
        _pageSetup.CenterFooter = companyName; // 公司名
        _pageSetup.RightFooter = "共&N页";   // 总页数
    }
    
    /// <summary>
    /// 设置自定义页眉
    /// </summary>
    public void SetCustomHeader(string left = "", string center = "", string right = "")
    {
        _pageSetup.LeftHeader = left;
        _pageSetup.CenterHeader = center;
        _pageSetup.RightHeader = right;
    }
    
    /// <summary>
    /// 设置自定义页脚
    /// </summary>
    public void SetCustomFooter(string left = "", string center = "", string right = "")
    {
        _pageSetup.LeftFooter = left;
        _pageSetup.CenterFooter = center;
        _pageSetup.RightFooter = right;
    }
    
    /// <summary>
    /// 设置奇偶页不同的页眉页脚
    /// </summary>
    public void EnableOddEvenPagesHeaderFooter()
    {
        _pageSetup.OddAndEvenPagesHeaderFooter = true;
    }
    
    /// <summary>
    /// 设置首页不同的页眉页脚
    /// </summary>
    public void EnableDifferentFirstPage()
    {
        _pageSetup.DifferentFirstPageHeaderFooter = true;
    }
    
    /// <summary>
    /// 获取标准页眉页脚代码
    /// </summary>
    public string GetHeaderFooterCode(HeaderFooterType type)
    {
        return _pageSetup.GetStandardHeaderFooterCode((int)type);
    }
}

public enum HeaderFooterType
{
    PageNumber = 1,        // 页码
    TotalPages = 2,        // 总页数
    Date = 3,              // 日期
    Time = 4,              // 时间
    FilePath = 5,          // 文件路径
    FileName = 6,          // 文件名
    SheetName = 7          // 工作表名
}
```

### 应用场景：企业文档页眉页脚设计

```csharp
public class CorporateHeaderFooterDesign
{
    private readonly HeaderFooterManager _headerFooterManager;
    
    public CorporateHeaderFooterDesign(IExcelPageSetup pageSetup)
    {
        _headerFooterManager = new HeaderFooterManager(pageSetup);
    }
    
    /// <summary>
    /// 设置财务报告页眉页脚
    /// </summary>
    public void SetFinancialReportHeaders(string reportTitle, string department)
    {
        // 启用奇偶页不同的页眉页脚
        _headerFooterManager.EnableOddEvenPagesHeaderFooter();
        _headerFooterManager.EnableDifferentFirstPage();
        
        // 首页页眉
        _headerFooterManager.SetCustomHeader(
            left: "",
            center: $"财务报告 - {reportTitle}",
            right: $"部门：{department}"
        );
        
        // 首页页脚
        _headerFooterManager.SetCustomFooter(
            left: "机密文件",
            center: "内部使用",
            right: "&D"
        );
    }
    
    /// <summary>
    /// 设置销售报告页眉页脚
    /// </summary>
    public void SetSalesReportHeaders(string period, string region)
    {
        _headerFooterManager.SetCustomHeader(
            left: $"销售报告 - {period}",
            center: $"区域：{region}",
            right: "第&P页"
        );
        
        _headerFooterManager.SetCustomFooter(
            left: "&F",
            center: "销售部门",
            right: "共&N页"
        );
    }
    
    /// <summary>
    /// 设置项目报告页眉页脚
    /// </summary>
    public void SetProjectReportHeaders(string projectName, string projectManager)
    {
        _headerFooterManager.SetCustomHeader(
            left: $"项目：{projectName}",
            center: $"项目经理：{projectManager}",
            right: "&D &T"  // 日期和时间
        );
        
        _headerFooterManager.SetCustomFooter(
            left: "项目文档",
            center: "版本1.0",
            right: "第&P页/共&N页"
        );
    }
}
```

## 打印选项控制

### 打印区域和标题设置

打印区域和标题设置是确保大型表格正确打印的关键功能。

```csharp
public class PrintOptionsManager
{
    private readonly IExcelPageSetup _pageSetup;
    
    public PrintOptionsManager(IExcelPageSetup pageSetup)
    {
        _pageSetup = pageSetup ?? throw new ArgumentNullException(nameof(pageSetup));
    }
    
    /// <summary>
    /// 设置打印区域
    /// </summary>
    public void SetPrintArea(string rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("打印区域地址不能为空");
            
        _pageSetup.PrintArea = rangeAddress;
    }
    
    /// <summary>
    /// 清除打印区域
    /// </summary>
    public void ClearPrintArea()
    {
        _pageSetup.PrintArea = "";
    }
    
    /// <summary>
    /// 设置打印标题行
    /// </summary>
    public void SetPrintTitleRows(string rowsAddress)
    {
        _pageSetup.PrintTitleRows = rowsAddress;
    }
    
    /// <summary>
    /// 设置打印标题列
    /// </summary>
    public void SetPrintTitleColumns(string columnsAddress)
    {
        _pageSetup.PrintTitleColumns = columnsAddress;
    }
    
    /// <summary>
    /// 设置网格线打印
    /// </summary>
    public void SetPrintGridlines(bool enabled)
    {
        _pageSetup.PrintGridlines = enabled;
    }
    
    /// <summary>
    /// 设置行列标题打印
    /// </summary>
    public void SetPrintHeadings(bool enabled)
    {
        _pageSetup.PrintHeadings = enabled;
    }
    
    /// <summary>
    /// 设置打印顺序
    /// </summary>
    public void SetPrintOrder(XlOrder order)
    {
        _pageSetup.Order = order;
    }
    
    /// <summary>
    /// 设置打印质量
    /// </summary>
    public void SetPrintQuality(object quality)
    {
        _pageSetup.PrintQuality = quality;
    }
    
    /// <summary>
    /// 设置草稿模式打印
    /// </summary>
    public void SetDraftMode(bool enabled)
    {
        _pageSetup.Draft = enabled;
    }
    
    /// <summary>
    /// 设置黑白打印
    /// </summary>
    public void SetBlackAndWhite(bool enabled)
    {
        _pageSetup.BlackAndWhite = enabled;
    }
}
```

### 应用场景：大型数据表打印设置

```csharp
public class LargeTablePrintSetup
{
    private readonly PrintOptionsManager _printOptionsManager;
    
    public LargeTablePrintSetup(IExcelPageSetup pageSetup)
    {
        _printOptionsManager = new PrintOptionsManager(pageSetup);
    }
    
    /// <summary>
    /// 设置销售数据表打印选项
    /// </summary>
    public void SetupSalesDataPrint(string dataRange, int titleRows = 1)
    {
        // 设置打印区域
        _printOptionsManager.SetPrintArea(dataRange);
        
        // 设置标题行（通常是表头）
        if (titleRows > 0)
        {
            _printOptionsManager.SetPrintTitleRows($"1:{titleRows}");
        }
        
        // 设置打印选项
        _printOptionsManager.SetPrintGridlines(true);  // 打印网格线
        _printOptionsManager.SetPrintHeadings(false);  // 不打印行列标题
        _printOptionsManager.SetPrintOrder(XlOrder.xlDownThenOver); // 先向下后向右
        _printOptionsManager.SetBlackAndWhite(false);  // 彩色打印
    }
    
    /// <summary>
    /// 设置库存报表打印选项
    /// </summary>
    public void SetupInventoryPrint(string dataRange, int titleRows = 2)
    {
        _printOptionsManager.SetPrintArea(dataRange);
        _printOptionsManager.SetPrintTitleRows($"1:{titleRows}");
        _printOptionsManager.SetPrintGridlines(true);
        _printOptionsManager.SetDraftMode(true); // 草稿模式，节省墨水
    }
    
    /// <summary>
    /// 设置财务报表打印选项
    /// </summary>
    public void SetupFinancialPrint(string dataRange, int titleRows = 3, int titleColumns = 1)
    {
        _printOptionsManager.SetPrintArea(dataRange);
        _printOptionsManager.SetPrintTitleRows($"1:{titleRows}");
        
        if (titleColumns > 0)
        {
            _printOptionsManager.SetPrintTitleColumns($"A:{GetColumnName(titleColumns)}");
        }
        
        _printOptionsManager.SetPrintGridlines(false); // 财务报表通常不打印网格线
        _printOptionsManager.SetBlackAndWhite(true);  // 黑白打印，更专业
    }
    
    private string GetColumnName(int columnNumber)
    {
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }
}
```

## 高级页面设置功能

### 分页符管理

分页符控制是多页文档打印的重要功能，确保内容在适当的位置分页。

```csharp
public class PageBreakManager
{
    private readonly IExcelWorksheet _worksheet;
    
    public PageBreakManager(IExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }
    
    /// <summary>
    /// 在指定行添加水平分页符
    /// </summary>
    public void AddHorizontalPageBreak(int row)
    {
        var pageBreak = _worksheet.HPageBreaks.Add(_worksheet.Rows[row]);
        // 可以进一步配置分页符属性
    }
    
    /// <summary>
    /// 在指定列添加垂直分页符
    /// </summary>
    public void AddVerticalPageBreak(int column)
    {
        var pageBreak = _worksheet.VPageBreaks.Add(_worksheet.Columns[column]);
    }
    
    /// <summary>
    /// 清除所有分页符
    /// </summary>
    public void ClearAllPageBreaks()
    {
        _worksheet.ResetAllPageBreaks();
    }
    
    /// <summary>
    /// 获取分页符信息
    /// </summary>
    public List<int> GetHorizontalPageBreaks()
    {
        var breaks = new List<int>();
        foreach (var pageBreak in _worksheet.HPageBreaks)
        {
            if (pageBreak.Location != null)
            {
                breaks.Add(pageBreak.Location.Row);
            }
        }
        return breaks;
    }
    
    /// <summary>
    /// 自动设置分页符基于内容
    /// </summary>
    public void AutoSetPageBreaks(int rowsPerPage = 40, int columnsPerPage = 8)
    {
        ClearAllPageBreaks();
        
        var usedRange = _worksheet.UsedRange;
        if (usedRange == null) return;
        
        int totalRows = usedRange.Rows.Count;
        int totalColumns = usedRange.Columns.Count;
        
        // 设置水平分页符
        for (int row = rowsPerPage; row < totalRows; row += rowsPerPage)
        {
            AddHorizontalPageBreak(row);
        }
        
        // 设置垂直分页符
        for (int col = columnsPerPage; col < totalColumns; col += columnsPerPage)
        {
            AddVerticalPageBreak(col);
        }
    }
}
```

### 应用场景：多页报表分页设置

```csharp
public class MultiPageReportSetup
{
    private readonly PageBreakManager _pageBreakManager;
    private readonly IExcelPageSetup _pageSetup;
    
    public MultiPageReportSetup(IExcelWorksheet worksheet, IExcelPageSetup pageSetup)
    {
        _pageBreakManager = new PageBreakManager(worksheet);
        _pageSetup = pageSetup;
    }
    
    /// <summary>
    /// 设置月度销售报告分页
    /// </summary>
    public void SetupMonthlySalesReportPaging()
    {
        // 每月数据分页
        _pageBreakManager.AutoSetPageBreaks(rowsPerPage: 35, columnsPerPage: 10);
        
        // 设置页面布局
        var layoutManager = new PageLayoutManager(_pageSetup);
        layoutManager.SetA4Landscape();
        layoutManager.SetFitToPage(1, 0); // 适合宽度
        
        // 设置页眉页脚
        var headerManager = new HeaderFooterManager(_pageSetup);
        headerManager.SetCustomHeader(
            left: "月度销售报告",
            center: DateTime.Now.ToString("yyyy年MM月"),
            right: "第&P页"
        );
    }
    
    /// <summary>
    /// 设置产品目录分页
    /// </summary>
    public void SetupProductCatalogPaging()
    {
        // 每个产品类别分页
        _pageBreakManager.AutoSetPageBreaks(rowsPerPage: 25, columnsPerPage: 6);
        
        var layoutManager = new PageLayoutManager(_pageSetup);
        layoutManager.SetA4Portrait();
        
        var headerManager = new HeaderFooterManager(_pageSetup);
        headerManager.SetCustomHeader(
            left: "产品目录",
            center: "",
            right: "&D"
        );
    }
}
```

## 打印预览和输出控制

### 打印预览设置

打印预览功能帮助用户在打印前确认文档布局。

```csharp
public class PrintPreviewManager
{
    private readonly IExcelApplication _excelApp;
    
    public PrintPreviewManager(IExcelApplication excelApp)
    {
        _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
    }
    
    /// <summary>
    /// 显示打印预览
    /// </summary>
    public void ShowPrintPreview()
    {
        _excelApp.ActiveSheet.PrintPreview();
    }
    
    /// <summary>
    /// 显示指定范围的打印预览
    /// </summary>
    public void ShowRangePrintPreview(string rangeAddress)
    {
        var range = _excelApp.ActiveSheet.Range(rangeAddress);
        range.PrintPreview();
    }
    
    /// <summary>
    /// 设置打印份数
    /// </summary>
    public void SetPrintCopies(int copies)
    {
        _excelApp.ActiveSheet.PageSetup.PrintCopies = copies;
    }
    
    /// <summary>
    /// 设置打印页码范围
    /// </summary>
    public void SetPrintRange(int fromPage, int toPage)
    {
        _excelApp.ActiveSheet.PageSetup.PrintPagesFrom = fromPage;
        _excelApp.ActiveSheet.PageSetup.PrintPagesTo = toPage;
    }
}
```

### 应用场景：批量打印管理系统

```csharp
public class BatchPrintManager
{
    private readonly PrintPreviewManager _previewManager;
    private readonly PrintOptionsManager _printOptionsManager;
    
    public BatchPrintManager(IExcelApplication excelApp, IExcelPageSetup pageSetup)
    {
        _previewManager = new PrintPreviewManager(excelApp);
        _printOptionsManager = new PrintOptionsManager(pageSetup);
    }
    
    /// <summary>
    /// 批量打印月度报告
    /// </summary>
    public void BatchPrintMonthlyReports(List<string> sheetNames, int copies = 1)
    {
        foreach (var sheetName in sheetNames)
        {
            var worksheet = _excelApp.Worksheets[sheetName];
            if (worksheet != null)
            {
                // 设置打印选项
                _printOptionsManager.SetPrintArea(worksheet.UsedRange.Address);
                _printOptionsManager.SetPrintTitleRows("1:2");
                
                // 设置打印份数
                _previewManager.SetPrintCopies(copies);
                
                // 显示预览（实际应用中可能直接打印）
                _previewManager.ShowPrintPreview();
            }
        }
    }
    
    /// <summary>
    /// 打印指定范围的页面
    /// </summary>
    public void PrintSpecificPages(int fromPage, int toPage, int copies = 1)
    {
        _previewManager.SetPrintRange(fromPage, toPage);
        _previewManager.SetPrintCopies(copies);
        _previewManager.ShowPrintPreview();
    }
}
```

## 性能优化和最佳实践

### 批量页面设置优化

```csharp
public class BatchPageSetupOptimizer
{
    /// <summary>
    /// 批量应用页面设置，减少COM调用
    /// </summary>
    public static void ApplyBatchPageSettings(IExcelPageSetup pageSetup, PageSettings settings)
    {
        // 禁用屏幕更新以提高性能
        using (var screenUpdater = new ScreenUpdater(pageSetup))
        {
            // 应用所有设置
            pageSetup.Orientation = settings.Orientation;
            pageSetup.PaperSize = settings.PaperSize;
            pageSetup.Zoom = settings.Zoom;
            
            // 页边距设置
            pageSetup.LeftMargin = settings.LeftMargin;
            pageSetup.RightMargin = settings.RightMargin;
            pageSetup.TopMargin = settings.TopMargin;
            pageSetup.BottomMargin = settings.BottomMargin;
            
            // 页眉页脚设置
            pageSetup.LeftHeader = settings.LeftHeader;
            pageSetup.CenterHeader = settings.CenterHeader;
            pageSetup.RightHeader = settings.RightHeader;
            
            // 应用设置
            pageSetup.Apply();
        }
    }
}

public class PageSettings
{
    public XlPageOrientation Orientation { get; set; } = XlPageOrientation.xlPortrait;
    public XlPaperSize PaperSize { get; set; } = XlPaperSize.xlPaperA4;
    public int Zoom { get; set; } = 100;
    public double LeftMargin { get; set; } = 1.0;
    public double RightMargin { get; set; } = 1.0;
    public double TopMargin { get; set; } = 1.0;
    public double BottomMargin { get; set; } = 1.0;
    public string LeftHeader { get; set; } = "";
    public string CenterHeader { get; set; } = "";
    public string RightHeader { get; set; } = "";
}

public class ScreenUpdater : IDisposable
{
    private readonly IExcelApplication _excelApp;
    private readonly bool _originalScreenUpdating;
    
    public ScreenUpdater(IExcelPageSetup pageSetup)
    {
        // 获取Excel应用程序实例
        _excelApp = GetExcelAppFromPageSetup(pageSetup);
        _originalScreenUpdating = _excelApp.ScreenUpdating;
        _excelApp.ScreenUpdating = false;
    }
    
    public void Dispose()
    {
        _excelApp.ScreenUpdating = _originalScreenUpdating;
    }
    
    private IExcelApplication GetExcelAppFromPageSetup(IExcelPageSetup pageSetup)
    {
        // 实现获取Excel应用程序实例的逻辑
        return null; // 简化示例
    }
}
```

## 总结

通过MudTools.OfficeInterop.Excel的页面布局和打印设置功能，开发者可以：

1. **精确控制页面布局**：方向、纸张大小、缩放比例
2. **专业页边距配置**：标准、窄、宽及自定义页边距
3. **丰富页眉页脚设计**：支持各种标准代码和自定义内容
4. **灵活打印选项**：打印区域、标题、网格线等控制
5. **智能分页管理**：自动和手动分页符设置
6. **批量打印支持**：多文档、多份数打印管理

这些功能结合实际业务场景，可以帮助创建专业、美观的打印文档，满足企业级应用的严格要求。