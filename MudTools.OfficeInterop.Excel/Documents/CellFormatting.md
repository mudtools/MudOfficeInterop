# 第七篇：单元格格式设置 - 打造专业美观的Excel报表

## 引言：Excel自动化的"美容师"

在Excel自动化开发中，如果说数据是报表的"骨架"，那么格式设置就是报表的"皮肤"和"妆容"！一个精心设计的报表不仅能够准确传达信息，更能通过美观的视觉效果提升专业形象，增强可读性。

想象一下这样的场景：你花费了大量时间整理数据、计算指标、分析趋势，最终生成了一份内容丰富的报表。但是，当这份报表呈现在领导面前时，却因为格式混乱、颜色搭配不当而显得不够专业。这就像是一位才华横溢的演员穿着不合身的衣服登台表演——内容再好，也难以给人留下深刻印象。

MudTools.OfficeInterop.Excel项目就像是专业的"报表美容师"，它提供了完整的格式设置接口，让开发者能够以编程方式精确控制Excel文档的每一个视觉细节。从字体选择到颜色搭配，从边框设计到填充效果，每一个细节都能得到完美的呈现。

本篇将带你探索Excel格式设置的奥秘，学习如何通过代码打造专业美观的报表。准备好让你的Excel自动化报表从"素颜"变成"精致妆容"了吗？

## 字体格式设置详解

### 基础字体属性设置

字体是Excel格式设置的基础，MudTools.OfficeInterop.Excel通过`IExcelFont`接口提供了全面的字体控制能力：

```csharp
using MudTools.OfficeInterop.Excel;
using System.Drawing;

public class FontFormatManager
{
    /// <summary>
    /// 设置基础字体格式
    /// </summary>
    public void SetBasicFontFormat(IExcelRange range)
    {
        // 获取字体对象
        var font = range.Font;
        
        // 设置字体名称和大小
        font.Name = "微软雅黑";
        font.Size = 12;
        
        // 设置字体样式
        font.Bold = true;      // 粗体
        font.Italic = false;   // 非斜体
        font.Underline = XlUnderlineStyle.xlUnderlineStyleSingle; // 单下划线
        
        // 设置字体颜色
        font.Color = Color.Blue; // 蓝色字体
        font.ColorIndex = 5;     // 使用颜色索引
    }
    
    /// <summary>
    /// 设置标题字体格式
    /// </summary>
    public void SetTitleFontFormat(IExcelRange titleRange)
    {
        var font = titleRange.Font;
        
        font.Name = "黑体";
        font.Size = 16;
        font.Bold = true;
        font.Color = Color.DarkBlue;
        
        // 设置特殊效果
        font.Strikethrough = false; // 无删除线
        font.Superscript = false;   // 非上标
        font.Subscript = false;     // 非下标
    }
    
    /// <summary>
    /// 设置数据字体格式
    /// </summary>
    public void SetDataFontFormat(IExcelRange dataRange)
    {
        var font = dataRange.Font;
        
        font.Name = "宋体";
        font.Size = 10;
        font.Bold = false;
        font.Color = Color.Black;
    }
}
```

### 高级字体应用场景

#### 场景1：财务报表字体设置

```csharp
public class FinancialReportFontManager
{
    public void FormatFinancialReport(IExcelWorksheet worksheet)
    {
        // 报表标题
        var titleRange = worksheet.Range("A1:F1");
        SetFinancialTitleFont(titleRange);
        
        // 表头
        var headerRange = worksheet.Range("A2:F2");
        SetFinancialHeaderFont(headerRange);
        
        // 数据区域
        var dataRange = worksheet.Range("A3:F20");
        SetFinancialDataFont(dataRange);
        
        // 总计行
        var totalRange = worksheet.Range("A21:F21");
        SetFinancialTotalFont(totalRange);
    }
    
    private void SetFinancialTitleFont(IExcelRange range)
    {
        var font = range.Font;
        font.Name = "黑体";
        font.Size = 18;
        font.Bold = true;
        font.Color = Color.DarkGreen;
        range.Merge(); // 合并单元格
        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
    }
    
    private void SetFinancialHeaderFont(IExcelRange range)
    {
        var font = range.Font;
        font.Name = "微软雅黑";
        font.Size = 11;
        font.Bold = true;
        font.Color = Color.White;
        
        // 设置背景色
        range.Interior.Color = Color.DarkBlue;
    }
    
    private void SetFinancialDataFont(IExcelRange range)
    {
        var font = range.Font;
        font.Name = "宋体";
        font.Size = 10;
        font.Bold = false;
        font.Color = Color.Black;
    }
    
    private void SetFinancialTotalFont(IExcelRange range)
    {
        var font = range.Font;
        font.Name = "微软雅黑";
        font.Size = 11;
        font.Bold = true;
        font.Color = Color.DarkRed;
        
        // 设置双下边框
        range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
    }
}
```

## 边框设置技术

### 基础边框设置

边框是表格结构的重要视觉元素，MudTools.OfficeInterop.Excel提供了精细的边框控制：

```csharp
public class BorderFormatManager
{
    /// <summary>
    /// 设置完整边框
    /// </summary>
    public void SetCompleteBorder(IExcelRange range)
    {
        var borders = range.Borders;
        
        // 设置所有边框
        borders.LineStyle = XlLineStyle.xlContinuous;
        borders.Color = Color.Black;
        borders.Weight = XlBorderWeight.xlThin;
    }
    
    /// <summary>
    /// 设置外边框
    /// </summary>
    public void SetOuterBorder(IExcelRange range)
    {
        var borders = range.Borders;
        
        // 设置外边框（粗线）
        borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeTop].Color = Color.Black;
        borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
        
        borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
        borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
        
        borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeLeft].Color = Color.Black;
        borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
        
        borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeRight].Color = Color.Black;
        borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
    }
    
    /// <summary>
    /// 设置内部网格线
    /// </summary>
    public void SetGridBorder(IExcelRange range)
    {
        var borders = range.Borders;
        
        // 设置内部网格线（细线）
        borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlInsideHorizontal].Color = Color.Gray;
        borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlHairline;
        
        borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlInsideVertical].Color = Color.Gray;
        borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlHairline;
    }
    
    /// <summary>
    /// 设置特殊边框样式
    /// </summary>
    public void SetSpecialBorder(IExcelRange range, BorderStyle style)
    {
        var borders = range.Borders;
        
        switch (style)
        {
            case BorderStyle.Dashed:
                borders.LineStyle = XlLineStyle.xlDash;
                borders.Color = Color.DarkGray;
                break;
                
            case BorderStyle.Dotted:
                borders.LineStyle = XlLineStyle.xlDot;
                borders.Color = Color.LightGray;
                break;
                
            case BorderStyle.Double:
                borders.LineStyle = XlLineStyle.xlDouble;
                borders.Color = Color.Black;
                borders.Weight = XlBorderWeight.xlThick;
                break;
        }
    }
}

public enum BorderStyle
{
    Dashed,
    Dotted,
    Double
}
```

### 高级边框应用

#### 场景2：数据表格边框设计

```csharp
public class DataTableBorderManager
{
    public void FormatDataTable(IExcelWorksheet worksheet, int startRow, int endRow, int columnCount)
    {
        // 获取数据区域
        var dataRange = worksheet.Range(
            worksheet.Cells[startRow, 1], 
            worksheet.Cells[endRow, columnCount]);
        
        // 设置表头边框
        var headerRange = worksheet.Range(
            worksheet.Cells[startRow, 1], 
            worksheet.Cells[startRow, columnCount]);
        SetHeaderBorder(headerRange);
        
        // 设置数据区域边框
        var bodyRange = worksheet.Range(
            worksheet.Cells[startRow + 1, 1], 
            worksheet.Cells[endRow, columnCount]);
        SetBodyBorder(bodyRange);
        
        // 设置总计行边框
        var totalRange = worksheet.Range(
            worksheet.Cells[endRow, 1], 
            worksheet.Cells[endRow, columnCount]);
        SetTotalBorder(totalRange);
    }
    
    private void SetHeaderBorder(IExcelRange range)
    {
        var borders = range.Borders;
        
        // 外边框 - 粗线
        borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
        borders[XlBordersIndex.xlEdgeTop].Color = Color.DarkBlue;
        
        borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
        borders[XlBordersIndex.xlEdgeBottom].Color = Color.DarkBlue;
        
        // 内部垂直分隔线
        borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlHairline;
        borders[XlBordersIndex.xlInsideVertical].Color = Color.LightGray;
    }
    
    private void SetBodyBorder(IExcelRange range)
    {
        var borders = range.Borders;
        
        // 所有边框 - 细线
        borders.LineStyle = XlLineStyle.xlContinuous;
        borders.Weight = XlBorderWeight.xlThin;
        borders.Color = Color.Gray;
        
        // 交替行背景色
        SetAlternatingRowColors(range);
    }
    
    private void SetTotalBorder(IExcelRange range)
    {
        var borders = range.Borders;
        
        // 上边框 - 双线（分隔数据与总计）
        borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDouble;
        borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;
        borders[XlBordersIndex.xlEdgeTop].Color = Color.DarkRed;
        
        // 下边框 - 粗线
        borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
        borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
        borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
    }
    
    private void SetAlternatingRowColors(IExcelRange range)
    {
        int rowCount = range.Rows.Count;
        
        for (int i = 1; i <= rowCount; i++)
        {
            var rowRange = range.Rows[i];
            
            if (i % 2 == 0) // 偶数行
            {
                rowRange.Interior.Color = Color.FromArgb(240, 240, 240); // 浅灰色
            }
            else // 奇数行
            {
                rowRange.Interior.Color = Color.White; // 白色
            }
        }
    }
}
```

## 填充和背景设置

### 单元格填充设置

单元格填充（背景色）是增强视觉效果的重要手段：

```csharp
public class FillFormatManager
{
    /// <summary>
    /// 设置纯色填充
    /// </summary>
    public void SetSolidFill(IExcelRange range, Color color)
    {
        var interior = range.Interior;
        
        interior.Color = color;
        interior.Pattern = XlPattern.xlPatternSolid;
        interior.PatternColorIndex = XlColorIndex.xlColorIndexAutomatic;
    }
    
    /// <summary>
    /// 设置渐变填充
    /// </summary>
    public void SetGradientFill(IExcelRange range, Color color1, Color color2)
    {
        var interior = range.Interior;
        
        interior.Pattern = XlPattern.xlPatternLinearGradient;
        
        // 设置渐变方向（从左到右）
        interior.Gradient.Degree = 0;
        interior.Gradient.ColorStops.Clear();
        
        // 添加颜色停止点
        var stop1 = interior.Gradient.ColorStops.Add(0);
        stop1.Color = color1;
        
        var stop2 = interior.Gradient.ColorStops.Add(1);
        stop2.Color = color2;
    }
    
    /// <summary>
    /// 设置图案填充
    /// </summary>
    public void SetPatternFill(IExcelRange range, XlPattern pattern, Color patternColor, Color backgroundColor)
    {
        var interior = range.Interior;
        
        interior.Pattern = pattern;
        interior.PatternColor = patternColor;
        interior.Color = backgroundColor;
    }
    
    /// <summary>
    /// 设置条件填充（基于值）
    /// </summary>
    public void SetConditionalFill(IExcelRange range, double threshold)
    {
        // 获取数据值
        var values = range.Value as object[,];
        
        if (values != null)
        {
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            
            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    var cell = range.Cells[i, j];
                    var value = values[i - 1, j - 1];
                    
                    if (value is double numericValue)
                    {
                        if (numericValue > threshold)
                        {
                            // 高于阈值 - 绿色
                            cell.Interior.Color = Color.LightGreen;
                        }
                        else if (numericValue < -threshold)
                        {
                            // 低于负阈值 - 浅红色
                            cell.Interior.Color = Color.LightCoral;
                        }
                        else
                        {
                            // 正常范围 - 白色
                            cell.Interior.Color = Color.White;
                        }
                    }
                }
            }
        }
    }
}
```

### 高级填充应用

#### 场景3：状态指示器填充

```csharp
public class StatusIndicatorManager
{
    public void FormatStatusIndicators(IExcelWorksheet worksheet, int statusColumn, int startRow, int endRow)
    {
        for (int row = startRow; row <= endRow; row++)
        {
            var statusCell = worksheet.Cells[row, statusColumn];
            var statusValue = statusCell.Value?.ToString() ?? "";
            
            SetStatusFill(statusCell, statusValue);
        }
    }
    
    private void SetStatusFill(IExcelRange cell, string status)
    {
        var interior = cell.Interior;
        
        switch (status.ToUpper())
        {
            case "完成":
            case "COMPLETED":
                interior.Color = Color.LightGreen;
                cell.Font.Color = Color.DarkGreen;
                break;
                
            case "进行中":
            case "IN PROGRESS":
                interior.Color = Color.LightYellow;
                cell.Font.Color = Color.DarkOrange;
                break;
                
            case "延期":
            case "DELAYED":
                interior.Color = Color.LightCoral;
                cell.Font.Color = Color.DarkRed;
                break;
                
            case "待开始":
            case "PENDING":
                interior.Color = Color.LightGray;
                cell.Font.Color = Color.DarkGray;
                break;
                
            default:
                interior.Color = Color.White;
                cell.Font.Color = Color.Black;
                break;
        }
        
        // 设置边框
        cell.Borders.LineStyle = XlLineStyle.xlContinuous;
        cell.Borders.Weight = XlBorderWeight.xlThin;
        cell.Borders.Color = Color.Gray;
        
        // 居中对齐
        cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
    }
}
```

## 数字格式设置

### 基础数字格式

数字格式是Excel格式设置的核心功能，直接影响数据的可读性：

```csharp
public class NumberFormatManager
{
    /// <summary>
    /// 设置货币格式
    /// </summary>
    public void SetCurrencyFormat(IExcelRange range)
    {
        range.NumberFormat = "¥#,##0.00"; // 人民币格式
        
        // 或者使用本地化格式
        range.NumberFormatLocal = "￥#,##0.00";
    }
    
    /// <summary>
    /// 设置百分比格式
    /// </summary>
    public void SetPercentageFormat(IExcelRange range, int decimalPlaces = 2)
    {
        string format = $"0.{new string('0', decimalPlaces)}%";
        range.NumberFormat = format;
    }
    
    /// <summary>
    /// 设置日期时间格式
    /// </summary>
    public void SetDateTimeFormat(IExcelRange range, DateTimeFormat format)
    {
        switch (format)
        {
            case DateTimeFormat.ShortDate:
                range.NumberFormat = "yyyy-mm-dd";
                break;
                
            case DateTimeFormat.LongDate:
                range.NumberFormat = "yyyy年mm月dd日";
                break;
                
            case DateTimeFormat.ShortDateTime:
                range.NumberFormat = "yyyy-mm-dd hh:mm";
                break;
                
            case DateTimeFormat.LongDateTime:
                range.NumberFormat = "yyyy年mm月dd日 hh时mm分";
                break;
        }
    }
    
    /// <summary>
    /// 设置自定义数字格式
    /// </summary>
    public void SetCustomFormat(IExcelRange range, CustomFormatType formatType)
    {
        string format = GetCustomFormatString(formatType);
        range.NumberFormat = format;
    }
    
    private string GetCustomFormatString(CustomFormatType formatType)
    {
        return formatType switch
        {
            CustomFormatType.PhoneNumber => "###-####-####",
            CustomFormatType.SocialSecurity => "###-##-####",
            CustomFormatType.ZipCode => "00000-0000",
            CustomFormatType.Scientific => "0.00E+00",
            CustomFormatType.Fraction => "# ?/?",
            _ => "General"
        };
    }
}

public enum DateTimeFormat
{
    ShortDate,
    LongDate,
    ShortDateTime,
    LongDateTime
}

public enum CustomFormatType
{
    PhoneNumber,
    SocialSecurity,
    ZipCode,
    Scientific,
    Fraction
}
```

### 高级数字格式应用

#### 场景4：财务报表数字格式

```csharp
public class FinancialNumberFormatManager
{
    public void FormatFinancialNumbers(IExcelWorksheet worksheet, FinancialReportType reportType)
    {
        switch (reportType)
        {
            case FinancialReportType.BalanceSheet:
                FormatBalanceSheet(worksheet);
                break;
                
            case FinancialReportType.IncomeStatement:
                FormatIncomeStatement(worksheet);
                break;
                
            case FinancialReportType.CashFlow:
                FormatCashFlowStatement(worksheet);
                break;
        }
    }
    
    private void FormatBalanceSheet(IExcelWorksheet worksheet)
    {
        // 资产类科目 - 正数格式
        var assetsRange = worksheet.Range("C5:C20");
        assetsRange.NumberFormat = "#,##0.00";
        
        // 负债和权益类科目 - 负数用括号表示
        var liabilitiesRange = worksheet.Range("D5:D20");
        liabilitiesRange.NumberFormat = "#,##0.00;(#,##0.00);-";
        
        // 比率指标 - 百分比
        var ratioRange = worksheet.Range("E5:E20");
        ratioRange.NumberFormat = "0.00%";
        
        // 总计行 - 加粗显示
        var totalRange = worksheet.Range("C21:D21");
        totalRange.NumberFormat = "#,##0.00_);(#,##0.00)";
        totalRange.Font.Bold = true;
    }
    
    private void FormatIncomeStatement(IExcelWorksheet worksheet)
    {
        // 收入类 - 正数格式
        var revenueRange = worksheet.Range("C5:C15");
        revenueRange.NumberFormat = "#,##0.00";
        
        // 成本费用类 - 负数用减号表示
        var expenseRange = worksheet.Range("D5:D15");
        expenseRange.NumberFormat = "#,##0.00;-#,##0.00;-";
        
        // 利润率 - 百分比带颜色
        var marginRange = worksheet.Range("E5:E15");
        marginRange.NumberFormat = "0.00%";
        
        // 设置条件格式：利润率低于10%显示红色
        SetMarginConditionalFormat(marginRange);
    }
    
    private void FormatCashFlowStatement(IExcelWorksheet worksheet)
    {
        // 现金流数据 - 带正负号
        var cashFlowRange = worksheet.Range("C5:C25");
        cashFlowRange.NumberFormat = "+#,##0.00;-#,##0.00;0.00";
        
        // 累计现金流 - 特殊格式
        var cumulativeRange = worksheet.Range("D5:D25");
        cumulativeRange.NumberFormat = "#,##0.00_);(#,##0.00)";
        
        // 现金流比率 - 百分比
        var cashRatioRange = worksheet.Range("E5:E25");
        cashRatioRange.NumberFormat = "0.00%";
    }
    
    private void SetMarginConditionalFormat(IExcelRange range)
    {
        // 这里使用条件格式API设置
        // 实际项目中会使用条件格式功能
        
        var values = range.Value as object[,];
        if (values != null)
        {
            int rows = values.GetLength(0);
            
            for (int i = 1; i <= rows; i++)
            {
                var cell = range.Cells[i, 1];
                var value = values[i - 1, 0];
                
                if (value is double margin && margin < 0.1) // 低于10%
                {
                    cell.Font.Color = Color.Red;
                    cell.Font.Bold = true;
                }
            }
        }
    }
}

public enum FinancialReportType
{
    BalanceSheet,
    IncomeStatement,
    CashFlow
}
```

## 对齐和文本方向设置

### 文本对齐设置

文本对齐影响内容的可读性和美观度：

```csharp
public class AlignmentManager
{
    /// <summary>
    /// 设置水平对齐
    /// </summary>
    public void SetHorizontalAlignment(IExcelRange range, XlHAlign alignment)
    {
        range.HorizontalAlignment = alignment;
    }
    
    /// <summary>
    /// 设置垂直对齐
    /// </summary>
    public void SetVerticalAlignment(IExcelRange range, XlVAlign alignment)
    {
        range.VerticalAlignment = alignment;
    }
    
    /// <summary>
    /// 设置文本方向
    /// </summary>
    public void SetTextOrientation(IExcelRange range, TextOrientation orientation)
    {
        switch (orientation)
        {
            case TextOrientation.Horizontal:
                range.Orientation = 0; // 水平
                break;
                
            case TextOrientation.Vertical:
                range.Orientation = 90; // 垂直
                break;
                
            case TextOrientation.Upward:
                range.Orientation = 45; // 向上倾斜
                break;
                
            case TextOrientation.Downward:
                range.Orientation = -45; // 向下倾斜
                break;
        }
    }
    
    /// <summary>
    /// 设置文本控制
    /// </summary>
    public void SetTextControl(IExcelRange range, bool wrapText, bool shrinkToFit)
    {
        range.WrapText = wrapText;
        range.ShrinkToFit = shrinkToFit;
    }
    
    /// <summary>
    /// 设置缩进
    /// </summary>
    public void SetIndent(IExcelRange range, int indentLevel)
    {
        range.IndentLevel = indentLevel;
    }
}

public enum TextOrientation
{
    Horizontal,
    Vertical,
    Upward,
    Downward
}
```

### 高级对齐应用

#### 场景5：多级标题对齐

```csharp
public class MultiLevelTitleManager
{
    public void FormatMultiLevelTitles(IExcelWorksheet worksheet, TitleLevel[] levels)
    {
        foreach (var level in levels)
        {
            var range = worksheet.Range(level.RangeAddress);
            FormatTitleLevel(range, level);
        }
    }
    
    private void FormatTitleLevel(IExcelRange range, TitleLevel level)
    {
        // 设置字体
        var font = range.Font;
        font.Name = level.FontName;
        font.Size = level.FontSize;
        font.Bold = level.IsBold;
        font.Color = level.FontColor;
        
        // 设置对齐
        range.HorizontalAlignment = level.HorizontalAlignment;
        range.VerticalAlignment = level.VerticalAlignment;
        
        // 设置缩进
        range.IndentLevel = level.IndentLevel;
        
        // 设置填充
        if (level.BackgroundColor != Color.Empty)
        {
            range.Interior.Color = level.BackgroundColor;
        }
        
        // 设置边框
        if (level.HasBorder)
        {
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
            range.Borders.Weight = XlBorderWeight.xlThin;
            range.Borders.Color = level.BorderColor;
        }
    }
}

public class TitleLevel
{
    public string RangeAddress { get; set; } = string.Empty;
    public string FontName { get; set; } = "微软雅黑";
    public double FontSize { get; set; } = 11;
    public bool IsBold { get; set; } = true;
    public Color FontColor { get; set; } = Color.Black;
    public XlHAlign HorizontalAlignment { get; set; } = XlHAlign.xlHAlignLeft;
    public XlVAlign VerticalAlignment { get; set; } = XlVAlign.xlVAlignCenter;
    public int IndentLevel { get; set; } = 0;
    public Color BackgroundColor { get; set; } = Color.Empty;
    public bool HasBorder { get; set; } = false;
    public Color BorderColor { get; set; } = Color.Gray;
}
```

## 样式管理和模板应用

### 样式对象管理

MudTools.OfficeInterop.Excel支持样式对象的管理和应用：

```csharp
public class StyleManager
{
    private IExcelWorkbook _workbook;
    
    public StyleManager(IExcelWorkbook workbook)
    {
        _workbook = workbook;
    }
    
    /// <summary>
    /// 创建自定义样式
    /// </summary>
    public IExcelStyle CreateCustomStyle(string styleName, StyleProperties properties)
    {
        var styles = _workbook.Styles;
        
        // 检查样式是否已存在
        var existingStyle = styles[styleName];
        if (existingStyle != null)
        {
            return existingStyle;
        }
        
        // 创建新样式
        var newStyle = styles.Add(styleName);
        
        // 设置样式属性
        if (properties.Font != null)
        {
            newStyle.Font.Name = properties.Font.Name;
            newStyle.Font.Size = properties.Font.Size;
            newStyle.Font.Bold = properties.Font.Bold;
            newStyle.Font.Color = properties.Font.Color;
        }
        
        if (properties.NumberFormat != null)
        {
            newStyle.NumberFormat = properties.NumberFormat;
        }
        
        // 设置边框
        if (properties.BorderStyle != null)
        {
            newStyle.Borders.LineStyle = properties.BorderStyle.Value;
        }
        
        // 设置填充
        if (properties.FillColor != Color.Empty)
        {
            newStyle.Interior.Color = properties.FillColor;
        }
        
        return newStyle;
    }
    
    /// <summary>
    /// 应用样式到范围
    /// </summary>
    public void ApplyStyle(IExcelRange range, string styleName)
    {
        var style = _workbook.Styles[styleName];
        if (style != null)
        {
            range.Style = styleName;
        }
    }
    
    /// <summary>
    /// 批量应用样式
    /// </summary>
    public void ApplyStylesToWorksheet(IExcelWorksheet worksheet, StyleMapping[] mappings)
    {
        foreach (var mapping in mappings)
        {
            var range = worksheet.Range(mapping.RangeAddress);
            ApplyStyle(range, mapping.StyleName);
        }
    }
    
    /// <summary>
    /// 基于现有样式创建变体
    /// </summary>
    public IExcelStyle CreateStyleVariant(string baseStyleName, string variantName, StyleModification modification)
    {
        var baseStyle = _workbook.Styles[baseStyleName];
        if (baseStyle == null) return null;
        
        var variantStyle = _workbook.Styles.Add(variantName);
        
        // 复制基础样式属性
        variantStyle.Font.Name = baseStyle.Font?.Name ?? "Calibri";
        variantStyle.Font.Size = modification.NewSize ?? (baseStyle.Font?.Size ?? 11);
        variantStyle.Font.Bold = modification.MakeBold ?? baseStyle.Font?.Bold ?? false;
        variantStyle.Font.Color = modification.NewColor ?? (baseStyle.Font?.Color ?? Color.Black);
        
        variantStyle.NumberFormat = modification.NewNumberFormat ?? baseStyle.NumberFormat;
        
        return variantStyle;
    }
}

public class StyleProperties
{
    public FontProperties Font { get; set; }
    public string NumberFormat { get; set; } = "General";
    public XlLineStyle? BorderStyle { get; set; }
    public Color FillColor { get; set; } = Color.Empty;
}

public class FontProperties
{
    public string Name { get; set; } = "Calibri";
    public double Size { get; set; } = 11;
    public bool Bold { get; set; } = false;
    public Color Color { get; set; } = Color.Black;
}

public class StyleMapping
{
    public string RangeAddress { get; set; } = string.Empty;
    public string StyleName { get; set; } = string.Empty;
}

public class StyleModification
{
    public double? NewSize { get; set; }
    public bool? MakeBold { get; set; }
    public Color? NewColor { get; set; }
    public string NewNumberFormat { get; set; }
}
```

## 实际应用案例

### 案例1：销售报表格式设置

```csharp
public class SalesReportFormatter
{
    public void FormatSalesReport(IExcelWorksheet worksheet, SalesData data)
    {
        // 1. 设置报表标题
        FormatReportTitle(worksheet);
        
        // 2. 设置表头格式
        FormatTableHeaders(worksheet);
        
        // 3. 设置数据区域格式
        FormatDataArea(worksheet, data);
        
        // 4. 设置汇总区域格式
        FormatSummaryArea(worksheet, data);
        
        // 5. 设置条件格式
        ApplyConditionalFormatting(worksheet, data);
    }
    
    private void FormatReportTitle(IExcelWorksheet worksheet)
    {
        var titleRange = worksheet.Range("A1:F1");
        titleRange.Merge();
        titleRange.Value = "月度销售报告";
        
        // 标题格式
        titleRange.Font.Name = "黑体";
        titleRange.Font.Size = 16;
        titleRange.Font.Bold = true;
        titleRange.Font.Color = Color.DarkBlue;
        titleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        titleRange.Interior.Color = Color.LightBlue;
    }
    
    private void FormatTableHeaders(IExcelWorksheet worksheet)
    {
        var headers = new[] { "产品名称", "销售数量", "单价", "销售额", "完成率", "状态" };
        
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = worksheet.Cells[3, i + 1];
            cell.Value = headers[i];
            
            // 表头格式
            cell.Font.Name = "微软雅黑";
            cell.Font.Size = 11;
            cell.Font.Bold = true;
            cell.Font.Color = Color.White;
            cell.Interior.Color = Color.DarkBlue;
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            
            // 设置边框
            cell.Borders.LineStyle = XlLineStyle.xlContinuous;
            cell.Borders.Weight = XlBorderWeight.xlThin;
            cell.Borders.Color = Color.White;
        }
    }
    
    private void FormatDataArea(IExcelWorksheet worksheet, SalesData data)
    {
        int startRow = 4;
        int endRow = startRow + data.Items.Count - 1;
        
        var dataRange = worksheet.Range(worksheet.Cells[startRow, 1], worksheet.Cells[endRow, 6]);
        
        // 设置数据格式
        dataRange.Font.Name = "宋体";
        dataRange.Font.Size = 10;
        
        // 设置数字格式
        worksheet.Range(worksheet.Cells[startRow, 2], worksheet.Cells[endRow, 2]).NumberFormat = "#,##0"; // 数量
        worksheet.Range(worksheet.Cells[startRow, 3], worksheet.Cells[endRow, 3]).NumberFormat = "¥#,##0.00"; // 单价
        worksheet.Range(worksheet.Cells[startRow, 4], worksheet.Cells[endRow, 4]).NumberFormat = "¥#,##0.00"; // 销售额
        worksheet.Range(worksheet.Cells[startRow, 5], worksheet.Cells[endRow, 5]).NumberFormat = "0.00%"; // 完成率
        
        // 设置交替行颜色
        for (int i = startRow; i <= endRow; i++)
        {
            var rowRange = worksheet.Range(worksheet.Cells[i, 1], worksheet.Cells[i, 6]);
            
            if (i % 2 == 0)
            {
                rowRange.Interior.Color = Color.FromArgb(245, 245, 245); // 浅灰色
            }
            else
            {
                rowRange.Interior.Color = Color.White;
            }
            
            // 设置边框
            rowRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders.Weight = XlBorderWeight.xlThin;
            rowRange.Borders.Color = Color.LightGray;
        }
    }
    
    private void FormatSummaryArea(IExcelWorksheet worksheet, SalesData data)
    {
        int summaryRow = data.Items.Count + 5;
        
        // 总计行
        var totalRange = worksheet.Range(worksheet.Cells[summaryRow, 1], worksheet.Cells[summaryRow, 6]);
        totalRange.Font.Bold = true;
        totalRange.Font.Color = Color.DarkRed;
        totalRange.Interior.Color = Color.LightYellow;
        
        // 设置双线上边框
        totalRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDouble;
        totalRange.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;
        totalRange.Borders[XlBordersIndex.xlEdgeTop].Color = Color.DarkRed;
    }
    
    private void ApplyConditionalFormatting(IExcelWorksheet worksheet, SalesData data)
    {
        // 根据完成率设置颜色
        int startRow = 4;
        
        for (int i = 0; i < data.Items.Count; i++)
        {
            var completionCell = worksheet.Cells[startRow + i, 5]; // 完成率列
            var statusCell = worksheet.Cells[startRow + i, 6]; // 状态列
            
            var completionRate = data.Items[i].CompletionRate;
            
            if (completionRate >= 1.0) // 超额完成
            {
                completionCell.Font.Color = Color.DarkGreen;
                statusCell.Value = "超额完成";
                statusCell.Font.Color = Color.DarkGreen;
            }
            else if (completionRate >= 0.8) // 正常完成
            {
                completionCell.Font.Color = Color.Blue;
                statusCell.Value = "正常完成";
                statusCell.Font.Color = Color.Blue;
            }
            else if (completionRate >= 0.6) // 基本完成
            {
                completionCell.Font.Color = Color.Orange;
                statusCell.Value = "基本完成";
                statusCell.Font.Color = Color.Orange;
            }
            else // 未完成
            {
                completionCell.Font.Color = Color.Red;
                statusCell.Value = "未完成";
                statusCell.Font.Color = Color.Red;
            }
        }
    }
}

public class SalesData
{
    public List<SalesItem> Items { get; set; } = new List<SalesItem>();
}

public class SalesItem
{
    public string ProductName { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal SalesAmount { get; set; }
    public double CompletionRate { get; set; }
}
```

## 性能优化建议

### 批量格式设置

```csharp
public class BatchFormatOptimizer
{
    /// <summary>
    /// 批量设置格式（性能优化版本）
    /// </summary>
    public void BatchFormatOptimized(IExcelWorksheet worksheet, FormatOperation[] operations)
    {
        // 禁用屏幕更新
        worksheet.Application.ScreenUpdating = false;
        
        try
        {
            // 按区域分组操作
            var groupedOperations = operations
                .GroupBy(op => op.RangeAddress)
                .ToList();
            
            foreach (var group in groupedOperations)
            {
                var range = worksheet.Range(group.Key);
                
                // 一次性应用所有格式设置
                ApplyMultipleFormats(range, group.ToArray());
            }
        }
        finally
        {
            // 恢复屏幕更新
            worksheet.Application.ScreenUpdating = true;
        }
    }
    
    private void ApplyMultipleFormats(IExcelRange range, FormatOperation[] operations)
    {
        foreach (var operation in operations)
        {
            switch (operation.Type)
            {
                case FormatType.Font:
                    ApplyFontFormat(range, operation.Parameters);
                    break;
                case FormatType.Border:
                    ApplyBorderFormat(range, operation.Parameters);
                    break;
                case FormatType.Fill:
                    ApplyFillFormat(range, operation.Parameters);
                    break;
                case FormatType.Number:
                    ApplyNumberFormat(range, operation.Parameters);
                    break;
            }
        }
    }
    
    private void ApplyFontFormat(IExcelRange range, Dictionary<string, object> parameters)
    {
        var font = range.Font;
        
        if (parameters.TryGetValue("Name", out var name)) font.Name = name.ToString();
        if (parameters.TryGetValue("Size", out var size)) font.Size = Convert.ToDouble(size);
        if (parameters.TryGetValue("Bold", out var bold)) font.Bold = Convert.ToBoolean(bold);
        if (parameters.TryGetValue("Color", out var color)) font.Color = (Color)color;
    }
    
    // 其他应用方法...
}

public class FormatOperation
{
    public string RangeAddress { get; set; } = string.Empty;
    public FormatType Type { get; set; }
    public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
}

public enum FormatType
{
    Font,
    Border,
    Fill,
    Number
}
```

## 总结

本篇详细介绍了MudTools.OfficeInterop.Excel项目中单元格格式设置的各个方面，包括字体、边框、填充、数字格式和对齐设置。通过丰富的代码示例和实际应用场景，展示了如何创建专业美观的Excel报表。

### 关键要点：

1. **字体设置** - 控制文本的外观和可读性
2. **边框设计** - 定义表格结构和视觉层次
3. **填充效果** - 增强视觉吸引力和信息传达
4. **数字格式** - 确保数据呈现的专业性和准确性
5. **对齐控制** - 优化内容的布局和可读性
6. **样式管理** - 实现格式的一致性和可维护性

### 最佳实践：

- 使用批量操作提高性能
- 建立样式模板确保一致性
- 合理使用条件格式突出关键信息
- 考虑用户体验和可访问性

通过掌握这些格式设置技术，开发者能够创建出既美观又实用的Excel报表，满足各种业务场景的需求。