# .NET驾驭Word之力：玩转文本与格式

在前面的文章中，我们已经了解了Word对象模型的核心组件，包括Application、Document和Range对象。掌握了这些基础知识后，我们现在可以进一步深入到文档内容的处理，特别是文本的插入和格式化操作。

本文将详细介绍如何使用MudTools.OfficeInterop.Word库来操作Word文档中的文本内容，包括多种插入文本的方法、字体格式设置和段落格式设置。最后，我们将通过一个实战示例，创建一个格式规范的商业信函模板，来综合运用所学知识。

## 3.1 插入文本的多种方式

在Word文档自动化处理中，插入文本是最基本也是最重要的操作之一。MudTools.OfficeInterop.Word提供了多种方式来插入文本，每种方式都有其适用场景。

### 使用 Range.Text 属性

[Range](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordRange.cs#L14-L467)对象是Word对象模型中最核心的组件之一，它代表文档中的一个连续区域。通过设置[Range.Text](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordRange.cs#L47-L47)属性，我们可以轻松地在指定位置插入或替换文本。

```csharp
// 获取文档的整个内容范围
var range = document.Content;

// 在文档末尾插入文本
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "这是通过Range.Text属性插入的文本。\n";

// 替换文档中的所有内容
range.Text = "这是替换后的全新内容。";
```

[Range.Text](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordRange.cs#L47-L47)属性是最直接的文本操作方式，适合于需要精确控制文本位置的场景。

#### 应用场景：动态报告生成

在企业环境中，经常需要根据数据动态生成报告。例如，财务部门需要每月生成财务报告，其中包含关键指标数据。

```csharp
using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;

// 财务数据模型
public class FinancialData
{
    public string Department { get; set; }
    public decimal Revenue { get; set; }
    public decimal Expenses { get; set; }
    public decimal Profit => Revenue - Expenses;
    public double GrowthRate { get; set; }
}

// 财务报告生成器
public class FinancialReportGenerator
{
    /// <summary>
    /// 生成财务报告
    /// </summary>
    /// <param name="data">财务数据列表</param>
    /// <param name="reportMonth">报告月份</param>
    public void GenerateFinancialReport(List<FinancialData> data, DateTime reportMonth)
    {
        try
        {
            // 使用模板创建报告文档
            using var wordApp = WordFactory.CreateFrom(@"C:\Templates\FinancialReportTemplate.dotx");
            var document = wordApp.ActiveDocument;
            
            // 隐藏Word应用程序以提高性能
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 替换报告标题中的月份信息
            document.FindAndReplace("[MONTH]", reportMonth.ToString("yyyy年MM月"));
            
            // 定位到数据表格位置
            var tableBookmark = document.Bookmarks["FinancialDataTable"];
            if (tableBookmark != null)
            {
                var tableRange = tableBookmark.Range;
                
                // 创建表格（标题行+数据行）
                var table = document.Tables.Add(tableRange, data.Count + 1, 5);
                
                // 设置表头
                table.Cell(1, 1).Range.Text = "部门";
                table.Cell(1, 2).Range.Text = "收入";
                table.Cell(1, 3).Range.Text = "支出";
                table.Cell(1, 4).Range.Text = "利润";
                table.Cell(1, 5).Range.Text = "增长率";
                
                // 填充数据
                for (int i = 0; i < data.Count; i++)
                {
                    var item = data[i];
                    table.Cell(i + 2, 1).Range.Text = item.Department;
                    table.Cell(i + 2, 2).Range.Text = item.Revenue.ToString("C");
                    table.Cell(i + 2, 3).Range.Text = item.Expenses.ToString("C");
                    table.Cell(i + 2, 4).Range.Text = item.Profit.ToString("C");
                    table.Cell(i + 2, 5).Range.Text = $"{item.GrowthRate:P2}";
                }
                
                // 格式化表格
                table.Borders.Enable = 1;
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        table.Cell(i, j).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }
            }
            
            // 保存报告
            string outputPath = $@"C:\Reports\FinancialReport_{reportMonth:yyyyMM}.docx";
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            document.Close();
            
            Console.WriteLine($"财务报告已生成: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成财务报告时发生错误: {ex.Message}");
        }
    }
}
```

### 使用 Selection 对象

[Selection](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordSelection.cs#L12-L293)对象代表文档中当前选中的区域。通过[Selection](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordSelection.cs#L12-L293)对象，我们可以像在Word界面中操作一样插入文本。

```csharp
// 获取当前选择区域
var selection = document.Selection;

// 插入文本
selection.InsertText("这是通过Selection对象插入的文本。");

// 插入段落
selection.InsertParagraph();

// 插入换行符
selection.InsertLineBreak();
```

虽然[Selection](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordSelection.cs#L12-L293)对象使用起来很直观，但在自动化处理中，我们通常不推荐将其作为主要方式，因为它依赖于当前光标位置，可能导致不可预期的结果。

#### 应用场景：交互式文档编辑器

在某些场景中，可能需要开发一个交互式文档编辑器，允许用户通过界面操作文档。

```csharp
// 交互式文档编辑器
public class InteractiveDocumentEditor
{
    private IWordApplication _wordApp;
    private IWordDocument _document;
    
    /// <summary>
    /// 初始化编辑器
    /// </summary>
    public void InitializeEditor()
    {
        try
        {
            // 创建可见的Word应用程序实例
            _wordApp = WordFactory.BlankWorkbook();
            _wordApp.Visibility = WordAppVisibility.Visible;
            _document = _wordApp.ActiveDocument;
            
            // 显示欢迎信息
            var selection = _document.Selection;
            selection.Font.Name = "微软雅黑";
            selection.Font.Size = 14;
            selection.Font.Bold = true;
            selection.Text = "欢迎使用交互式文档编辑器\n\n";
            
            selection.Font.Bold = false;
            selection.Font.Size = 12;
            selection.Text = "请开始编辑您的文档...\n";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"初始化编辑器时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 在光标位置插入文本
    /// </summary>
    /// <param name="text">要插入的文本</param>
    public void InsertTextAtCursor(string text)
    {
        try
        {
            if (_document != null)
            {
                var selection = _document.Selection;
                selection.InsertText(text);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"插入文本时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 在光标位置插入日期
    /// </summary>
    public void InsertCurrentDate()
    {
        try
        {
            if (_document != null)
            {
                var selection = _document.Selection;
                selection.InsertText(DateTime.Now.ToString("yyyy年MM月dd日"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"插入日期时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 清理资源
    /// </summary>
    public void Cleanup()
    {
        try
        {
            _document?.Close(false); // 不保存更改
            _wordApp?.Quit();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"清理资源时发生错误: {ex.Message}");
        }
    }
}
```

### 使用 Document.Content 和 Document.Paragraphs 等集合

通过文档的集合属性，我们可以更结构化地操作文档内容。

```csharp
// 使用Document.Content获取整个文档内容
var contentRange = document.Content;
contentRange.Text += "添加到文档末尾的内容。\n";

// 使用Paragraphs集合添加新段落
var newParagraph = document.Paragraphs.Add();
newParagraph.Range.Text = "这是一个新段落。";
```

这种方式适合于需要按结构化方式处理文档内容的场景。

#### 应用场景：批量文档处理

在企业环境中，经常需要批量处理大量文档，例如为多份合同添加相同的条款。

```csharp
// 批量文档处理器
public class BatchDocumentProcessor
{
    /// <summary>
    /// 为多个文档添加通用条款
    /// </summary>
    /// <param name="documentPaths">文档路径列表</param>
    /// <param name="termsText">通用条款文本</param>
    public void AddTermsToDocuments(List<string> documentPaths, string termsText)
    {
        foreach (var documentPath in documentPaths)
        {
            try
            {
                // 打开文档
                using var wordApp = WordFactory.Open(documentPath);
                var document = wordApp.ActiveDocument;
                
                // 隐藏Word应用程序以提高性能
                wordApp.Visibility = WordAppVisibility.Hidden;
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                
                // 在文档末尾添加通用条款
                var contentRange = document.Content;
                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                
                // 添加分页符
                contentRange.InsertBreak(WdBreakType.wdPageBreak);
                
                // 添加条款标题
                contentRange.Text += "\n通用条款\n\n";
                contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                contentRange.Font.Bold = true;
                contentRange.Font.Size = 14;
                
                // 重置格式
                contentRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                contentRange.Font.Bold = false;
                contentRange.Font.Size = 12;
                contentRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                
                // 添加条款内容
                contentRange.Text += termsText;
                
                // 保存文档
                document.Save();
                document.Close();
                
                Console.WriteLine($"已为文档添加通用条款: {documentPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文档 {documentPath} 时发生错误: {ex.Message}");
            }
        }
    }
    
    /// <summary>
    /// 为多个文档添加页眉和页脚
    /// </summary>
    /// <param name="documentPaths">文档路径列表</param>
    /// <param name="headerText">页眉文本</param>
    /// <param name="footerText">页脚文本</param>
    public void AddHeaderFooterToDocuments(List<string> documentPaths, string headerText, string footerText)
    {
        foreach (var documentPath in documentPaths)
        {
            try
            {
                // 打开文档
                using var wordApp = WordFactory.Open(documentPath);
                var document = wordApp.ActiveDocument;
                
                // 隐藏Word应用程序以提高性能
                wordApp.Visibility = WordAppVisibility.Hidden;
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                
                // 设置页眉
                foreach (Section section in document.Sections)
                {
                    var headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Text = headerText;
                    headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                
                // 设置页脚
                foreach (Section section in document.Sections)
                {
                    var footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Text = footerText;
                    footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                
                // 保存文档
                document.Save();
                document.Close();
                
                Console.WriteLine($"已为文档添加页眉页脚: {documentPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文档 {documentPath} 时发生错误: {ex.Message}");
            }
        }
    }
}
```

## 3.2 字体格式设置 (Font Object)

在文档处理中，字体格式设置是提升文档可读性和美观度的重要手段。MudTools.OfficeInterop.Word通过[IWordFont](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Content/Text/IWordFont.cs#L12-L78)接口提供了丰富的字体格式设置功能。

### 基本字体属性设置

[IWordFont](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Content/Text/IWordFont.cs#L12-L78)接口提供了设置字体名称、大小、颜色、加粗、斜体、下划线等基本属性的方法。

```csharp
// 获取文档内容范围
var range = document.Content;

// 设置字体名称
range.Font.Name = "微软雅黑";

// 设置字体大小（单位：磅）
range.Font.Size = 12;

// 设置字体颜色
range.Font.Color = WdColor.wdColorBlue;

// 设置加粗
range.Font.Bold = true;

// 设置斜体
range.Font.Italic = true;

// 设置下划线
range.Font.Underline = true;
```

### 高级字体属性设置

除了基本属性外，[IWordFont](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Content/Text/IWordFont.cs#L12-L78)还支持更多高级属性设置：

```csharp
// 设置上标
range.Font.Superscript = true;

// 设置下标
range.Font.Subscript = true;

// 设置字符间距
range.Font.Spacing = 2; // 增加2磅间距

// 设置字符缩放比例
range.Font.Scaling = 150; // 150%大小

// 设置字符位置偏移
range.Font.Position = 3; // 上移3磅
```

#### 应用场景：科学文档格式化

在学术或科研环境中，经常需要处理包含数学公式、化学方程式等特殊格式的文档。

```csharp
// 科学文档格式化器
public class ScientificDocumentFormatter
{
    /// <summary>
    /// 格式化化学方程式
    /// </summary>
    /// <param name="document">Word文档</param>
    public void FormatChemicalEquations(IWordDocument document)
    {
        try
        {
            // 查找所有化学方程式（假设用[chem]标记）
            var range = document.Content.Duplicate;
            
            while (range.FindAndReplace("[chem]", "", WdReplace.wdReplaceNone))
            {
                // 获取方程式内容
                var equationRange = document.Range(range.Start, range.End);
                
                // 格式化为下标
                equationRange.Font.Subscript = true;
                equationRange.Font.Size = 10;
                equationRange.Font.Name = "Cambria Math";
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"格式化化学方程式时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 格式化数学公式
    /// </summary>
    /// <param name="document">Word文档</param>
    public void FormatMathematicalFormulas(IWordDocument document)
    {
        try
        {
            // 查找所有数学公式（假设用[math]标记）
            var range = document.Content.Duplicate;
            
            while (range.FindAndReplace("[math]", "", WdReplace.wdReplaceNone))
            {
                // 获取公式内容
                var formulaRange = document.Range(range.Start, range.End);
                
                // 设置字体为数学字体
                formulaRange.Font.Name = "Cambria Math";
                formulaRange.Font.Size = 12;
                
                // 处理上标（用^标记）
                var superscriptRange = formulaRange.Duplicate;
                while (superscriptRange.FindAndReplace("^", "", WdReplace.wdReplaceNone))
                {
                    var supRange = document.Range(superscriptRange.Start, superscriptRange.End + 1);
                    supRange.Font.Superscript = true;
                    supRange.Font.Size = 8;
                }
                
                // 处理下标（用_标记）
                var subscriptRange = formulaRange.Duplicate;
                while (subscriptRange.FindAndReplace("_", "", WdReplace.wdReplaceNone))
                {
                    var subRange = document.Range(subscriptRange.Start, subscriptRange.End + 1);
                    subRange.Font.Subscript = true;
                    subRange.Font.Size = 8;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"格式化数学公式时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 格式化代码片段
    /// </summary>
    /// <param name="document">Word文档</param>
    public void FormatCodeSnippets(IWordDocument document)
    {
        try
        {
            // 查找所有代码片段（假设用[code]标记）
            var range = document.Content.Duplicate;
            
            while (range.FindAndReplace("[code]", "", WdReplace.wdReplaceNone))
            {
                // 获取代码内容
                var codeRange = document.Range(range.Start, range.End);
                
                // 设置等宽字体
                codeRange.Font.Name = "Consolas";
                codeRange.Font.Size = 10;
                codeRange.Font.Bold = false;
                
                // 设置背景色
                codeRange.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                
                // 添加边框
                codeRange.Borders.Enable = 1;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"格式化代码片段时发生错误: {ex.Message}");
        }
    }
}
```

## 3.3 段落格式设置 (ParagraphFormat Object)

段落格式决定了文本的布局和视觉效果。通过[IWordParagraphFormat](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Formatting/Format/IWordParagraphFormat.cs#L12-L118)接口，我们可以设置段落的对齐方式、缩进、行距等属性。

### 段落对齐方式

```csharp
// 获取文档内容范围
var range = document.Content;

// 设置段落对齐方式
range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 居中对齐
// 其他选项包括：
// WdParagraphAlignment.wdAlignParagraphLeft    - 左对齐
// WdParagraphAlignment.wdAlignParagraphRight   - 右对齐
// WdParagraphAlignment.wdAlignParagraphJustify - 两端对齐
```

### 缩进设置

```csharp
// 设置首行缩进（单位：磅）
range.ParagraphFormat.FirstLineIndent = 21; // 约等于2个字符宽度

// 设置左缩进
range.ParagraphFormat.LeftIndent = 36; // 约等于3个字符宽度

// 设置右缩进
range.ParagraphFormat.RightIndent = 18; // 约等于1.5个字符宽度
```

### 行距和间距设置

```csharp
// 设置行距规则
range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble; // 双倍行距
// 其他选项包括：
// WdLineSpacing.wdLineSpaceSingle     - 单倍行距
// WdLineSpacing.wdLineSpace1pt5       - 1.5倍行距
// WdLineSpacing.wdLineSpaceExactly    - 固定值行距
// WdLineSpacing.wdLineSpaceMultiple   - 多倍行距

// 设置段前间距（单位：磅）
range.ParagraphFormat.SpaceBefore = 12;

// 设置段后间距（单位：磅）
range.ParagraphFormat.SpaceAfter = 12;
```

#### 应用场景：文档样式统一化

在企业环境中，为了保持文档风格的一致性，经常需要对文档进行样式统一化处理。

```csharp
// 文档样式统一化工具
public class DocumentStyleUnifier
{
    /// <summary>
    /// 统一文档标题样式
    /// </summary>
    /// <param name="document">Word文档</param>
    public void UnifyHeadingStyles(IWordDocument document)
    {
        try
        {
            // 处理一级标题（以#开头的段落）
            var heading1Range = document.Content.Duplicate;
            while (heading1Range.FindAndReplace("# ", "", WdReplace.wdReplaceNone))
            {
                // 获取标题段落
                var para = heading1Range.Paragraphs.First();
                var paraRange = para.Range;
                
                // 设置一级标题样式
                paraRange.Font.Name = "黑体";
                paraRange.Font.Size = 16;
                paraRange.Font.Bold = true;
                paraRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paraRange.ParagraphFormat.SpaceBefore = 18;
                paraRange.ParagraphFormat.SpaceAfter = 12;
                
                // 移除标记符号
                paraRange.Text = paraRange.Text.Replace("# ", "");
            }
            
            // 处理二级标题（以##开头的段落）
            var heading2Range = document.Content.Duplicate;
            while (heading2Range.FindAndReplace("## ", "", WdReplace.wdReplaceNone))
            {
                // 获取标题段落
                var para = heading2Range.Paragraphs.First();
                var paraRange = para.Range;
                
                // 设置二级标题样式
                paraRange.Font.Name = "微软雅黑";
                paraRange.Font.Size = 14;
                paraRange.Font.Bold = true;
                paraRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paraRange.ParagraphFormat.SpaceBefore = 12;
                paraRange.ParagraphFormat.SpaceAfter = 6;
                
                // 移除标记符号
                paraRange.Text = paraRange.Text.Replace("## ", "");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"统一标题样式时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 统一文档正文样式
    /// </summary>
    /// <param name="document">Word文档</param>
    public void UnifyBodyTextStyles(IWordDocument document)
    {
        try
        {
            // 获取文档正文范围
            var bodyRange = document.Content;
            
            // 设置正文样式
            bodyRange.Font.Name = "仿宋_GB2312";
            bodyRange.Font.Size = 12;
            bodyRange.ParagraphFormat.FirstLineIndent = 28; // 首行缩进2字符
            bodyRange.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5; // 1.5倍行距
            bodyRange.ParagraphFormat.SpaceBefore = 0;
            bodyRange.ParagraphFormat.SpaceAfter = 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"统一正文样式时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 统一列表样式
    /// </summary>
    /// <param name="document">Word文档</param>
    public void UnifyListStyles(IWordDocument document)
    {
        try
        {
            // 处理无序列表（以-开头的行）
            var bulletRange = document.Content.Duplicate;
            while (bulletRange.FindAndReplace("- ", "", WdReplace.wdReplaceNone))
            {
                var listRange = document.Range(bulletRange.Start, bulletRange.End);
                
                // 应用项目符号列表格式
                listRange.ListFormat.ApplyBulletDefault();
                
                // 设置列表项格式
                listRange.ParagraphFormat.LeftIndent = 36;
                listRange.ParagraphFormat.FirstLineIndent = -18;
            }
            
            // 处理有序列表（以数字.开头的行）
            for (int i = 1; i <= 9; i++)
            {
                var numberedRange = document.Content.Duplicate;
                while (numberedRange.FindAndReplace($"{i}. ", "", WdReplace.wdReplaceNone))
                {
                    var listRange = document.Range(numberedRange.Start, numberedRange.End);
                    
                    // 应用编号列表格式
                    listRange.ListFormat.ApplyNumberDefault();
                    
                    // 设置列表项格式
                    listRange.ParagraphFormat.LeftIndent = 36;
                    listRange.ParagraphFormat.FirstLineIndent = -18;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"统一列表样式时发生错误: {ex.Message}");
        }
    }
}
```

## 3.4 实战：创建一个格式规范的商业信函模板

现在，让我们综合运用前面学到的知识，创建一个格式规范的商业信函模板。

```csharp
using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System;

public class BusinessLetterTemplate
{
    public void CreateBusinessLetter()
    {
        try
        {
            // 创建一个新的空白文档
            using var wordApp = WordFactory.BlankWorkbook();
            var document = wordApp.ActiveDocument;
            
            // 隐藏Word应用程序以提高性能
            wordApp.Visibility = WordAppVisibility.Hidden;
            
            // 禁止显示警告
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 设置文档整体字体
            document.Content.Font.Name = "仿宋_GB2312";
            document.Content.Font.Size = 12;
            
            // 插入发信人信息（右对齐）
            var senderRange = document.Range(0, 0);
            senderRange.Text = "发信人公司名称\n地址\n电话：XXX-XXXXXXX\n邮箱：xxxx@xxxx.com\n\n";
            senderRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            
            // 插入日期（右对齐）
            var dateRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            dateRange.Text = DateTime.Now.ToString("yyyy年MM月dd日") + "\n\n";
            dateRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            
            // 插入收信人信息（左对齐）
            var recipientRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            recipientRange.Text = "收信人姓名\n收信人职位\n收信人公司名称\n收信人地址\n\n";
            recipientRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            
            // 插入信件正文标题（居中，加粗）
            var titleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            titleRange.Text = "商务合作邀请函\n\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;
            titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 重置字体大小
            titleRange.Font.Size = 12;
            
            // 插入正文内容（首行缩进）
            var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            string content = "尊敬的合作伙伴：\n\n" +
                            "    首先感谢您一直以来对我们公司的关注与支持。我们诚挚地邀请您参与我们的新项目合作。" +
                            "该项目旨在通过双方的优势资源整合，实现互利共赢的目标。\n\n" +
                            "    我们相信，通过双方的精诚合作，必将开创更加美好的未来。期待您的积极回应，" +
                            "并希望能尽快与您展开深入的交流与探讨。\n\n" +
                            "    如有任何疑问，请随时与我们联系。\n\n" +
                            "此致\n敬礼！\n\n\n";
            contentRange.Text = content;
            
            // 设置正文段落格式（首行缩进2字符）
            contentRange.ParagraphFormat.FirstLineIndent = 28; // 约等于2个字符宽度
            
            // 插入发信人签名
            var signatureRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            signatureRange.Text = "发信人姓名\n发信人职位\n发信人公司名称";
            signatureRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            
            // 保存文档
            string outputPath = @"C:\Temp\BusinessLetterTemplate.docx";
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            
            Console.WriteLine($"商业信函模板已创建: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建商业信函模板时发生错误: {ex.Message}");
        }
    }
    
    public void CreateFormattedBusinessLetter(string senderCompany, string senderAddress, 
                                              string senderPhone, string senderEmail,
                                              string recipientName, string recipientTitle,
                                              string recipientCompany, string recipientAddress,
                                              string letterSubject, string letterContent)
    {
        try
        {
            // 基于模板创建文档
            using var wordApp = WordFactory.BlankWorkbook();
            var document = wordApp.ActiveDocument;
            
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 设置文档整体字体
            document.Content.Font.Name = "仿宋_GB2312";
            document.Content.Font.Size = 12;
            
            // 插入发信人信息
            var senderRange = document.Range(0, 0);
            senderRange.Text = $"{senderCompany}\n{senderAddress}\n电话：{senderPhone}\n邮箱：{senderEmail}\n\n";
            senderRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            
            // 插入日期
            var dateRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            dateRange.Text = DateTime.Now.ToString("yyyy年MM月dd日") + "\n\n";
            dateRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            
            // 插入收信人信息
            var recipientRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            recipientRange.Text = $"{recipientName}\n{recipientTitle}\n{recipientCompany}\n{recipientAddress}\n\n";
            recipientRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            
            // 插入信件标题
            var titleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            titleRange.Text = $"{letterSubject}\n\n";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 16;
            titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            titleRange.Font.Size = 12;
            
            // 插入正文内容
            var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            contentRange.Text = letterContent;
            contentRange.ParagraphFormat.FirstLineIndent = 28;
            
            // 插入发信人签名占位符
            var signatureRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            signatureRange.Text = "\n\n发信人签名：___________\n发信人姓名\n发信人职位\n发信人公司名称";
            signatureRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            
            // 保存文档
            string outputPath = $@"C:\Temp\{letterSubject}.docx";
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            
            Console.WriteLine($"格式化的商业信函已创建: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建格式化的商业信函时发生错误: {ex.Message}");
        }
    }
}
```

### 应用场景：企业文档自动化系统

基于我们学到的知识，可以构建一个完整的企业文档自动化系统：

```csharp
// 企业文档自动化系统
public class EnterpriseDocumentAutomationSystem
{
    /// <summary>
    /// 商务信函生成服务
    /// </summary>
    public class BusinessLetterService
    {
        /// <summary>
        /// 生成商务信函
        /// </summary>
        /// <param name="request">信函请求参数</param>
        /// <returns>生成的文档路径</returns>
        public string GenerateBusinessLetter(BusinessLetterRequest request)
        {
            try
            {
                using var wordApp = WordFactory.BlankWorkbook();
                var document = wordApp.ActiveDocument;
                
                wordApp.Visibility = WordAppVisibility.Hidden;
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                
                // 应用标准模板样式
                ApplyStandardStyles(document);
                
                // 填充信函内容
                FillLetterContent(document, request);
                
                // 保存文档
                string outputPath = $@"C:\Documents\BusinessLetters\{request.LetterId}.docx";
                document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
                document.Close();
                
                return outputPath;
            }
            catch (Exception ex)
            {
                throw new DocumentGenerationException($"生成商务信函时发生错误: {ex.Message}", ex);
            }
        }
        
        private void ApplyStandardStyles(IWordDocument document)
        {
            // 设置全局字体
            document.Content.Font.Name = "仿宋_GB2312";
            document.Content.Font.Size = 12;
            
            // 设置页面边距
            document.PageSetup.TopMargin = 72;    // 1英寸
            document.PageSetup.BottomMargin = 72; // 1英寸
            document.PageSetup.LeftMargin = 90;   // 1.25英寸
            document.PageSetup.RightMargin = 90;  // 1.25英寸
        }
        
        private void FillLetterContent(IWordDocument document, BusinessLetterRequest request)
        {
            // 插入发信人信息
            InsertSenderInfo(document, request.SenderInfo);
            
            // 插入日期
            InsertDate(document, request.LetterDate);
            
            // 插入收信人信息
            InsertRecipientInfo(document, request.RecipientInfo);
            
            // 插入信函标题
            InsertLetterTitle(document, request.Subject);
            
            // 插入正文内容
            InsertLetterContent(document, request.Content);
            
            // 插入签名
            InsertSignature(document, request.SenderInfo);
        }
        
        private void InsertSenderInfo(IWordDocument document, SenderInfo senderInfo)
        {
            var range = document.Range(0, 0);
            range.Text = $"{senderInfo.Company}\n{senderInfo.Address}\n电话：{senderInfo.Phone}\n邮箱：{senderInfo.Email}\n\n";
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }
        
        private void InsertDate(IWordDocument document, DateTime date)
        {
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = date.ToString("yyyy年MM月dd日") + "\n\n";
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }
        
        private void InsertRecipientInfo(IWordDocument document, RecipientInfo recipientInfo)
        {
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = $"{recipientInfo.Name}\n{recipientInfo.Title}\n{recipientInfo.Company}\n{recipientInfo.Address}\n\n";
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }
        
        private void InsertLetterTitle(IWordDocument document, string subject)
        {
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = $"{subject}\n\n";
            range.Font.Bold = true;
            range.Font.Size = 16;
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            range.Font.Size = 12;
        }
        
        private void InsertLetterContent(IWordDocument document, string content)
        {
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = content;
            range.ParagraphFormat.FirstLineIndent = 28;
        }
        
        private void InsertSignature(IWordDocument document, SenderInfo senderInfo)
        {
            var range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = $"\n\n发信人签名：___________\n{senderInfo.Name}\n{senderInfo.Title}\n{senderInfo.Company}";
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }
    }
    
    // 数据模型类
    public class BusinessLetterRequest
    {
        public string LetterId { get; set; }
        public SenderInfo SenderInfo { get; set; }
        public RecipientInfo RecipientInfo { get; set; }
        public DateTime LetterDate { get; set; }
        public string Subject { get; set; }
        public string Content { get; set; }
    }
    
    public class SenderInfo
    {
        public string Company { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string Name { get; set; }
        public string Title { get; set; }
    }
    
    public class RecipientInfo
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Company { get; set; }
        public string Address { get; set; }
    }
    
    public class DocumentGenerationException : Exception
    {
        public DocumentGenerationException(string message, Exception innerException) 
            : base(message, innerException)
        {
        }
    }
}
```

## 小结

本文详细介绍了使用MudTools.OfficeInterop.Word库操作Word文档文本和格式的方法：

1. **文本插入方式**：介绍了通过Range.Text属性、Selection对象和文档集合属性等多种方式插入文本，并提供了相应的应用场景和代码示例
2. **字体格式设置**：演示了如何使用IWordFont接口设置字体名称、大小、颜色、加粗、斜体等属性，并展示了在科学文档格式化中的应用
3. **段落格式设置**：展示了如何使用IWordParagraphFormat接口设置段落对齐方式、缩进、行距等属性，并提供了文档样式统一化的实际应用
4. **实战应用**：通过创建商业信函模板和企业文档自动化系统，综合运用了所学的文本和格式操作知识

掌握这些文本和格式操作技巧，可以帮助我们创建更加专业和美观的Word文档，为后续的文档自动化处理奠定坚实基础。

在下一篇文章中，我们将探讨更高级的主题，包括表格操作、图片插入和文档样式设置等内容。敬请期待！