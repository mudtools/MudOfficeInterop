# .NET驾驭Word之力：结构化文档元素操作

在前几篇文章中，我们学习了Word对象模型的基础知识、文本操作与格式设置等内容。掌握了这些基础知识后，我们现在可以进一步深入到文档的结构化元素操作，包括段落与节的管理、表格的创建与操作以及图片的插入等。

本文将详细介绍如何使用MudTools.OfficeInterop.Word库来操作Word文档中的结构化元素，包括段落与节的使用、表格的自动化操作以及图片与形状的插入。最后，我们将通过一个实战示例——创建一个包含多种结构化元素的员工信息表，来综合运用所学知识。

## 4.1 使用段落(Paragraphs)与节(Sections)

段落和节是Word文档中重要的结构化元素。段落用于组织文本内容，而节则用于对文档进行分段，以便为不同部分设置不同的页面布局。

### 遍历文档中的所有段落

在处理Word文档时，经常需要遍历文档中的所有段落以进行批量操作。通过[Paragraphs](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordDocument.cs#L317-L317)属性，我们可以轻松访问文档中的所有段落。

```csharp
using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

// 打开现有文档
using var wordApp = WordFactory.Open(@"C:\Documents\SampleDocument.docx");
var document = wordApp.ActiveDocument;

// 遍历文档中的所有段落
foreach (var paragraph in document.Paragraphs)
{
    // 输出段落文本
    Console.WriteLine(paragraph.GetText());
    
    // 为每个段落设置12磅的段后间距
    paragraph.SpaceAfter = 12;
    
    // 为每个段落设置1.5倍行距
    paragraph.LineSpacingRule = WdLineSpacing.wdLineSpace15;
}

// 或者通过索引访问特定段落
for (int i = 1; i <= document.ParagraphCount; i++)
{
    var paragraph = document.Paragraphs[i];
    // 处理段落内容
    Console.WriteLine($"第{i}段: {paragraph.GetText()}");
}
```

在上面的示例中，我们展示了两种遍历段落的方式：使用foreach循环和通过索引访问。每种方式都有其适用场景，foreach循环适用于需要处理所有段落的情况，而索引访问适用于需要精确控制处理顺序或只处理特定段落的情况。

#### 应用场景：文档格式标准化

在企业环境中，经常需要对大量文档进行格式标准化处理。例如，确保所有文档的段落间距、行距、字体等符合公司规范。

```csharp
/// <summary>
/// 文档格式标准化工具
/// </summary>
public class DocumentFormatter
{
    /// <summary>
    /// 标准化文档格式
    /// </summary>
    /// <param name="documentPath">文档路径</param>
    public void StandardizeDocument(string documentPath)
    {
        try
        {
            // 打开文档
            using var wordApp = WordFactory.Open(documentPath);
            var document = wordApp.ActiveDocument;
            
            // 隐藏Word应用程序以提高性能
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 遍历所有段落并标准化格式
            foreach (var paragraph in document.Paragraphs)
            {
                // 设置段落格式
                paragraph.SpaceAfter = 12;  // 段后间距12磅
                paragraph.SpaceBefore = 0;  // 段前间距0磅
                paragraph.LineSpacingRule = WdLineSpacing.wdLineSpace15; // 1.5倍行距
                
                // 设置字体格式
                paragraph.Range.Font.Name = "微软雅黑";
                paragraph.Range.Font.Size = 10.5f;
                
                // 设置对齐方式
                paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphJustify; // 两端对齐
            }
            
            // 保存文档
            document.Save();
            document.Close();
            
            Console.WriteLine($"文档 {documentPath} 格式标准化完成");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"格式标准化过程中发生错误: {ex.Message}");
        }
    }
}
```

### 使用节(Section)为文档的不同部分设置不同的页面布局

节是Word文档中用于分隔具有不同页面布局设置的区域。通过节，我们可以为文档的不同部分设置不同的页眉页脚、纸张方向、页边距等。

```csharp
// 添加新节并设置不同的页面方向
var sections = document.Sections;

// 获取当前节的数量
int sectionCount = sections.Count;

// 在文档末尾添加分节符以创建新节
document.AddSectionBreak(document.Content.End - 1, (int)WdSectionBreakType.wdSectionBreakNextPage);

// 获取新添加的节
var newSection = sections[sectionCount + 1];

// 为新节设置横向页面
newSection.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

// 为不同节设置不同的页眉
var firstSectionHeader = sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
firstSectionHeader.Range.Text = "这是第一节的页眉";

var newSectionHeader = newSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
newSectionHeader.Range.Text = "这是新节的页眉";
```

通过以上代码，我们可以为文档的不同部分设置不同的页面布局。这对于制作包含多种内容类型的复杂文档非常有用，例如在同一篇文档中既有纵向的文字说明，又有横向的表格数据。

#### 应用场景：制作混合布局报告

在制作技术报告或商业文档时，经常需要在同一篇文档中包含不同类型的页面布局。例如，文档正文使用纵向布局，而数据表格使用横向布局。

```csharp
/// <summary>
/// 混合布局报告生成器
/// </summary>
public class MixedLayoutReportGenerator
{
    /// <summary>
    /// 生成混合布局报告
    /// </summary>
    /// <param name="templatePath">模板路径</param>
    /// <param name="outputPath">输出路径</param>
    public void GenerateReport(string templatePath, string outputPath)
    {
        try
        {
            // 基于模板创建文档
            using var wordApp = WordFactory.CreateFrom(templatePath);
            var document = wordApp.ActiveDocument;
            
            // 隐藏Word应用程序
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 在文档末尾添加分节符，创建新节用于横向表格
            document.AddSectionBreak(document.Content.End - 1, 
                (int)WdSectionBreakType.wdSectionBreakNextPage);
            
            // 获取新节
            var dataSection = document.Sections[document.Sections.Count];
            
            // 设置新节为横向布局
            dataSection.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            
            // 在新节中添加标题
            var range = dataSection.Range;
            range.Collapse(WdCollapseDirection.wdCollapseStart);
            range.Text = "数据汇总表\n";
            range.Font.Bold = 1;
            range.Font.Size = 14;
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 添加表格
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            var table = document.Tables.Add(range, 10, 6); // 10行6列的表格
            
            // 填充表格数据
            PopulateTableData(table);
            
            // 保存文档
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            document.Close();
            
            Console.WriteLine($"混合布局报告已生成: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成报告时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 填充表格数据
    /// </summary>
    /// <param name="table">表格对象</param>
    private void PopulateTableData(IWordTable table)
    {
        // 表头
        string[] headers = { "序号", "产品名称", "销售数量", "单价", "总金额", "备注" };
        for (int i = 0; i < headers.Length; i++)
        {
            table.Cell(1, i + 1).Range.Text = headers[i];
            table.Cell(1, i + 1).Range.Font.Bold = 1;
            table.Cell(1, i + 1).VerticalAlignment = 
                WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        }
        
        // 示例数据
        string[,] data = {
            {"1", "产品A", "100", "50.00", "5000.00", ""},
            {"2", "产品B", "200", "30.00", "6000.00", ""},
            {"3", "产品C", "150", "40.00", "6000.00", ""},
            {"4", "产品D", "80", "70.00", "5600.00", ""},
            {"5", "产品E", "120", "35.00", "4200.00", ""}
        };
        
        // 填充数据
        for (int i = 0; i < data.GetLength(0); i++)
        {
            for (int j = 0; j < data.GetLength(1); j++)
            {
                table.Cell(i + 2, j + 1).Range.Text = data[i, j];
                table.Cell(i + 2, j + 1).VerticalAlignment = 
                    WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }
        }
        
        // 设置表格样式
        table.Borders.Enable = 1;
        table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
        table.PreferredWidth = 100;
    }
}
```

## 4.2 表格(Table)的自动化

表格是Word文档中用于组织和展示数据的重要元素。MudTools.OfficeInterop.Word库提供了丰富的API来创建、操作和格式化表格。

### 创建指定行数列的表格

使用[Tables.Add](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop.Word/Core/IWordTables.cs#L32-L38)方法，我们可以轻松地在文档中创建指定行列数的表格。

```csharp
// 在文档末尾创建一个5行4列的表格
var range = document.Content;
range.Collapse(WdCollapseDirection.wdCollapseEnd); // 将范围折叠到末尾

var table = document.Tables.Add(range, 5, 4);

// 设置表格标题行
table.Rows[1].Cells[1].Range.Text = "姓名";
table.Rows[1].Cells[2].Range.Text = "部门";
table.Rows[1].Cells[3].Range.Text = "职位";
table.Rows[1].Cells[4].Range.Text = "入职日期";

// 填充表格数据
string[,] employeeData = {
    {"张三", "技术部", "软件工程师", "2022-01-15"},
    {"李四", "市场部", "市场专员", "2021-11-20"},
    {"王五", "人事部", "人事经理", "2020-05-10"},
    {"赵六", "财务部", "会计师", "2022-03-08"}
};

for (int i = 0; i < employeeData.GetLength(0); i++)
{
    for (int j = 0; j < employeeData.GetLength(1); j++)
    {
        table.Cell(i + 2, j + 1).Range.Text = employeeData[i, j];
    }
}

// 设置标题行为粗体
for (int i = 1; i <= 4; i++)
{
    table.Cell(1, i).Range.Font.Bold = 1;
}
```

### 遍历单元格、写入数据、设置表格样式和边框

创建表格后，我们需要填充数据并设置样式。

```csharp
// 设置表格样式
table.TableStyle = "网格型";

// 设置表格边框
table.Borders.Enable = 1;
table.Borders.LineStyle = WdLineStyle.wdLineStyleSingle;
table.Borders.LineWidth = WdLineWidth.wdLineWidth150pt;

// 遍历所有单元格并设置对齐方式
foreach (var row in table.Rows)
{
    foreach (var cell in row.Cells)
    {
        cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
    }
}

// 设置表格宽度
table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
table.PreferredWidth = 100;
```

### 单元格合并与拆分

在实际应用中，我们经常需要合并或拆分单元格以满足不同的布局需求。

```csharp
// 合并单元格示例：合并第一行的所有单元格作为标题
var firstRow = table.Rows[1];
firstRow.Cells[1].Merge(firstRow.Cells[4]);

// 在合并后的单元格中添加标题文本
firstRow.Cells[1].Range.Text = "员工信息表";
firstRow.Cells[1].Range.Font.Bold = 1;
firstRow.Cells[1].Range.Font.Size = 14;
firstRow.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

// 拆分单元格示例：拆分指定单元格
var cellToSplit = table.Cell(3, 2);
cellToSplit.Split(2, 1); // 拆分为2行1列
```

#### 应用场景：自动化生成财务报表

在财务部门，经常需要生成各种财务报表，这些报表通常包含复杂的表格结构。通过自动化生成，可以大大提高工作效率并减少错误。

```csharp
/// <summary>
/// 财务报表生成器
/// </summary>
public class FinancialReportGenerator
{
    /// <summary>
    /// 生成财务报表
    /// </summary>
    /// <param name="financialData">财务数据</param>
    /// <param name="outputPath">输出路径</param>
    public void GenerateFinancialReport(FinancialData financialData, string outputPath)
    {
        try
        {
            // 创建新文档
            using var wordApp = WordFactory.BlankWorkbook();
            var document = wordApp.ActiveDocument;
            
            // 隐藏Word应用程序
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 添加标题
            AddReportTitle(document, "年度财务报告");
            
            // 添加报告基本信息
            AddReportInfo(document, financialData);
            
            // 添加收入明细表
            AddIncomeStatement(document, financialData.IncomeItems);
            
            // 添加资产负债表
            AddBalanceSheet(document, financialData.Assets, financialData.Liabilities, 
                financialData.Equity);
            
            // 保存文档
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            document.Close();
            
            Console.WriteLine($"财务报告已生成: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成财务报告时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 添加报告标题
    /// </summary>
    /// <param name="document">文档对象</param>
    /// <param name="title">标题文本</param>
    private void AddReportTitle(IWordDocument document, string title)
    {
        var titleParagraph = document.AddParagraph(0, title);
        titleParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        titleParagraph.Range.Font.Name = "微软雅黑";
        titleParagraph.Range.Font.Size = 18;
        titleParagraph.Range.Font.Bold = 1;
        
        // 添加空行
        document.AddParagraph(document.Content.End - 1);
    }
    
    /// <summary>
    /// 添加报告基本信息
    /// </summary>
    /// <param name="document">文档对象</param>
    /// <param name="financialData">财务数据</param>
    private void AddReportInfo(IWordDocument document, FinancialData financialData)
    {
        var infoParagraph = document.AddParagraph(document.Content.End - 1, 
            $"报告期间: {financialData.Period}\n" +
            $"编制单位: {financialData.CompanyName}\n" +
            $"货币单位: 人民币元\n");
        infoParagraph.Range.Font.Name = "微软雅黑";
        infoParagraph.Range.Font.Size = 12;
        
        // 添加空行
        document.AddParagraph(document.Content.End - 1);
    }
    
    /// <summary>
    /// 添加收入明细表
    /// </summary>
    /// <param name="document">文档对象</param>
    /// <param name="incomeItems">收入项目</param>
    private void AddIncomeStatement(IWordDocument document, List<IncomeItem> incomeItems)
    {
        // 添加表标题
        var titleParagraph = document.AddParagraph(document.Content.End - 1, "一、收入明细表");
        titleParagraph.Range.Font.Bold = 1;
        titleParagraph.Range.Font.Size = 14;
        
        // 添加空行
        document.AddParagraph(document.Content.End - 1);
        
        // 创建表格
        var range = document.Content;
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        var table = document.Tables.Add(range, incomeItems.Count + 2, 4); // 数据行+标题行+合计行
        
        // 设置表头
        string[] headers = { "项目", "本期金额", "上期金额", "增减率(%)" };
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = table.Cell(1, i + 1);
            cell.Range.Text = headers[i];
            cell.Range.Font.Bold = 1;
            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }
        
        // 填充数据
        decimal totalCurrent = 0, totalPrevious = 0;
        for (int i = 0; i < incomeItems.Count; i++)
        {
            var item = incomeItems[i];
            table.Cell(i + 2, 1).Range.Text = item.Name;
            table.Cell(i + 2, 2).Range.Text = item.CurrentAmount.ToString("N2");
            table.Cell(i + 2, 3).Range.Text = item.PreviousAmount.ToString("N2");
            table.Cell(i + 2, 4).Range.Text = item.ChangeRate.ToString("F2");
            
            totalCurrent += item.CurrentAmount;
            totalPrevious += item.PreviousAmount;
            
            // 设置对齐方式
            for (int j = 1; j <= 4; j++)
            {
                table.Cell(i + 2, j).VerticalAlignment = 
                    WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                if (j > 1)
                {
                    table.Cell(i + 2, j).Range.ParagraphFormat.Alignment = 
                        WdParagraphAlignment.wdAlignParagraphRight;
                }
            }
        }
        
        // 添加合计行
        table.Cell(incomeItems.Count + 2, 1).Range.Text = "合计";
        table.Cell(incomeItems.Count + 2, 1).Range.Font.Bold = 1;
        table.Cell(incomeItems.Count + 2, 2).Range.Text = totalCurrent.ToString("N2");
        table.Cell(incomeItems.Count + 2, 2).Range.Font.Bold = 1;
        table.Cell(incomeItems.Count + 2, 3).Range.Text = totalPrevious.ToString("N2");
        table.Cell(incomeItems.Count + 2, 3).Range.Font.Bold = 1;
        
        var totalChangeRate = totalPrevious != 0 ? 
            (totalCurrent - totalPrevious) / totalPrevious * 100 : 0;
        table.Cell(incomeItems.Count + 2, 4).Range.Text = totalChangeRate.ToString("F2");
        table.Cell(incomeItems.Count + 2, 4).Range.Font.Bold = 1;
        
        // 设置合计行对齐方式
        for (int j = 1; j <= 4; j++)
        {
            table.Cell(incomeItems.Count + 2, j).VerticalAlignment = 
                WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            if (j > 1)
            {
                table.Cell(incomeItems.Count + 2, j).Range.ParagraphFormat.Alignment = 
                    WdParagraphAlignment.wdAlignParagraphRight;
            }
        }
        
        // 设置表格样式
        table.Borders.Enable = 1;
        table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
        table.PreferredWidth = 100;
        
        // 添加空行
        document.AddParagraph(document.Content.End - 1);
        document.AddParagraph(document.Content.End - 1);
    }
    
    /// <summary>
    /// 添加资产负债表
    /// </summary>
    /// <param name="document">文档对象</param>
    /// <param name="assets">资产项目</param>
    /// <param name="liabilities">负债项目</param>
    /// <param name="equity">权益项目</param>
    private void AddBalanceSheet(IWordDocument document, List<AssetItem> assets, 
        List<LiabilityItem> liabilities, List<EquityItem> equity)
    {
        // 添加表标题
        var titleParagraph = document.AddParagraph(document.Content.End - 1, "二、资产负债表");
        titleParagraph.Range.Font.Bold = 1;
        titleParagraph.Range.Font.Size = 14;
        
        // 添加空行
        document.AddParagraph(document.Content.End - 1);
        
        // 创建表格
        var range = document.Content;
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        var table = document.Tables.Add(range, 
            Math.Max(assets.Count, liabilities.Count + equity.Count) + 1, 4);
        
        // 设置表头
        table.Cell(1, 1).Range.Text = "资产";
        table.Cell(1, 1).Range.Font.Bold = 1;
        table.Cell(1, 2).Range.Text = "金额";
        table.Cell(1, 2).Range.Font.Bold = 1;
        table.Cell(1, 3).Range.Text = "负债和权益";
        table.Cell(1, 3).Range.Font.Bold = 1;
        table.Cell(1, 4).Range.Text = "金额";
        table.Cell(1, 4).Range.Font.Bold = 1;
        
        // 设置表头格式
        for (int i = 1; i <= 4; i++)
        {
            table.Cell(1, i).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(1, i).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }
        
        // 填充资产数据
        for (int i = 0; i < assets.Count; i++)
        {
            table.Cell(i + 2, 1).Range.Text = assets[i].Name;
            table.Cell(i + 2, 2).Range.Text = assets[i].Amount.ToString("N2");
            
            // 设置格式
            table.Cell(i + 2, 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(i + 2, 2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(i + 2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }
        
        // 填充负债和权益数据
        int liabilityStartRow = 2;
        for (int i = 0; i < liabilities.Count; i++)
        {
            table.Cell(i + liabilityStartRow, 3).Range.Text = liabilities[i].Name;
            table.Cell(i + liabilityStartRow, 4).Range.Text = liabilities[i].Amount.ToString("N2");
            
            // 设置格式
            table.Cell(i + liabilityStartRow, 3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(i + liabilityStartRow, 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(i + liabilityStartRow, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }
        
        // 填充权益数据
        int equityStartRow = liabilityStartRow + liabilities.Count;
        for (int i = 0; i < equity.Count; i++)
        {
            table.Cell(i + equityStartRow, 3).Range.Text = equity[i].Name;
            table.Cell(i + equityStartRow, 4).Range.Text = equity[i].Amount.ToString("N2");
            
            // 设置格式
            table.Cell(i + equityStartRow, 3).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(i + equityStartRow, 4).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Cell(i + equityStartRow, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }
        
        // 计算资产合计
        decimal totalAssets = assets.Sum(a => a.Amount);
        table.Cell(assets.Count + 2, 1).Range.Text = "资产总计";
        table.Cell(assets.Count + 2, 1).Range.Font.Bold = 1;
        table.Cell(assets.Count + 2, 2).Range.Text = totalAssets.ToString("N2");
        table.Cell(assets.Count + 2, 2).Range.Font.Bold = 1;
        
        // 计算负债和权益合计
        decimal totalLiabilities = liabilities.Sum(l => l.Amount);
        decimal totalEquity = equity.Sum(e => e.Amount);
        table.Cell(Math.Max(assets.Count, liabilities.Count + equity.Count) + 1, 3).Range.Text = "负债和权益总计";
        table.Cell(Math.Max(assets.Count, liabilities.Count + equity.Count) + 1, 3).Range.Font.Bold = 1;
        table.Cell(Math.Max(assets.Count, liabilities.Count + equity.Count) + 1, 4).Range.Text = 
            (totalLiabilities + totalEquity).ToString("N2");
        table.Cell(Math.Max(assets.Count, liabilities.Count + equity.Count) + 1, 4).Range.Font.Bold = 1;
        
        // 设置表格样式
        table.Borders.Enable = 1;
        table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
        table.PreferredWidth = 100;
    }
}

/// <summary>
/// 财务数据模型
/// </summary>
public class FinancialData
{
    /// <summary>
    /// 报告期间
    /// </summary>
    public string Period { get; set; }
    
    /// <summary>
    /// 公司名称
    /// </summary>
    public string CompanyName { get; set; }
    
    /// <summary>
    /// 收入项目列表
    /// </summary>
    public List<IncomeItem> IncomeItems { get; set; }
    
    /// <summary>
    /// 资产项目列表
    /// </summary>
    public List<AssetItem> Assets { get; set; }
    
    /// <summary>
    /// 负债项目列表
    /// </summary>
    public List<LiabilityItem> Liabilities { get; set; }
    
    /// <summary>
    /// 权益项目列表
    /// </summary>
    public List<EquityItem> Equity { get; set; }
}

/// <summary>
/// 收入项目
/// </summary>
public class IncomeItem
{
    /// <summary>
    /// 项目名称
    /// </summary>
    public string Name { get; set; }
    
    /// <summary>
    /// 本期金额
    /// </summary>
    public decimal CurrentAmount { get; set; }
    
    /// <summary>
    /// 上期金额
    /// </summary>
    public decimal PreviousAmount { get; set; }
    
    /// <summary>
    /// 增减率
    /// </summary>
    public decimal ChangeRate => PreviousAmount != 0 ? 
        (CurrentAmount - PreviousAmount) / PreviousAmount * 100 : 0;
}

/// <summary>
/// 资产项目
/// </summary>
public class AssetItem
{
    /// <summary>
    /// 项目名称
    /// </summary>
    public string Name { get; set; }
    
    /// <summary>
    /// 金额
    /// </summary>
    public decimal Amount { get; set; }
}

/// <summary>
/// 负债项目
/// </summary>
public class LiabilityItem
{
    /// <summary>
    /// 项目名称
    /// </summary>
    public string Name { get; set; }
    
    /// <summary>
    /// 金额
    /// </summary>
    public decimal Amount { get; set; }
}

/// <summary>
/// 权益项目
/// </summary>
public class EquityItem
{
    /// <summary>
    /// 项目名称
    /// </summary>
    public string Name { get; set; }
    
    /// <summary>
    /// 金额
    /// </summary>
    public decimal Amount { get; set; }
}
```

## 4.3 图片与形状的插入

图片和形状能够丰富文档的视觉效果，使其更加生动和易于理解。Word提供了两种类型的图形对象：内嵌形状和浮动形状。

### 使用InlineShapes.AddPicture方法插入图片

内嵌形状是嵌入在文本行中的对象，它们随着文本移动而移动。

```csharp
// 在文档末尾插入内嵌图片
var range = document.Content;
range.Collapse(WdCollapseDirection.wdCollapseEnd);

// 使用InlineShapes.AddPicture方法插入图片
var inlineShape = document.InlineShapes.AddPicture(
    fileName: @"C:\Images\company_logo.png",
    linkToFile: false,        // 不链接到文件
    saveWithDocument: true    // 与文档一起保存
);

// 设置图片大小
inlineShape.Width = 100;
inlineShape.Height = 50;

// 添加图片说明文字
range.InsertAfter("\n公司Logo\n");
```

### 使用Shapes.AddPicture方法插入浮动图片并设置环绕方式

浮动形状是独立于文本流的对象，可以放置在页面上的任意位置，并可以设置文字环绕方式。

```csharp
// 插入浮动图片
var shape = document.Shapes.AddPicture(
    fileName: @"C:\Images\decorative_image.png",
    linkToFile: false,
    saveWithDocument: true,
    left: 300,    // 距离页面左边距300磅
    top: 150,     // 距离页面上边距150磅
    width: 150,
    height: 100
);

// 设置图片的环绕方式
shape.WrapFormat.Type = WdWrapType.wdWrapSquare;
shape.WrapFormat.DistanceTop = 10;
shape.WrapFormat.DistanceBottom = 10;
shape.WrapFormat.DistanceLeft = 10;
shape.WrapFormat.DistanceRight = 10;

// 设置图片的相对位置
shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
```

#### 应用场景：自动化制作产品宣传册

在市场营销领域，经常需要制作产品宣传册。通过自动化生成，可以快速制作大量标准化的宣传材料。

```csharp
/// <summary>
/// 产品宣传册生成器
/// </summary>
public class ProductBrochureGenerator
{
    /// <summary>
    /// 生成产品宣传册
    /// </summary>
    /// <param name="products">产品列表</param>
    /// <param name="templatePath">模板路径</param>
    /// <param name="outputPath">输出路径</param>
    public void GenerateBrochure(List<Product> products, string templatePath, string outputPath)
    {
        try
        {
            // 基于模板创建文档
            using var wordApp = WordFactory.CreateFrom(templatePath);
            var document = wordApp.ActiveDocument;
            
            // 隐藏Word应用程序
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            
            // 为每个产品添加页面
            foreach (var product in products)
            {
                AddProductPage(document, product);
            }
            
            // 保存文档
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            document.Close();
            
            Console.WriteLine($"产品宣传册已生成: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成产品宣传册时发生错误: {ex.Message}");
        }
    }
    
    /// <summary>
    /// 添加产品页面
    /// </summary>
    /// <param name="document">文档对象</param>
    /// <param name="product">产品信息</param>
    private void AddProductPage(IWordDocument document, Product product)
    {
        // 添加分页符
        document.AddPageBreak(document.Content.End - 1);
        
        // 添加产品名称
        var titleParagraph = document.AddParagraph(document.Content.End - 1, product.Name);
        titleParagraph.Range.Font.Name = "微软雅黑";
        titleParagraph.Range.Font.Size = 24;
        titleParagraph.Range.Font.Bold = 1;
        titleParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        titleParagraph.SpaceAfter = 20;
        
        // 添加产品图片
        if (File.Exists(product.ImagePath))
        {
            var range = document.Content;
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            var inlineShape = document.InlineShapes.AddPicture(
                fileName: product.ImagePath,
                linkToFile: false,
                saveWithDocument: true
            );
            
            // 设置图片大小（保持纵横比）
            inlineShape.LockAspectRatio = true;
            if (inlineShape.Width > 300)
            {
                inlineShape.Width = 300;
            }
            
            // 居中对齐图片
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 添加空行
            document.AddParagraph(document.Content.End - 1);
        }
        
        // 添加产品描述
        var descriptionParagraph = document.AddParagraph(document.Content.End - 1, product.Description);
        descriptionParagraph.Range.Font.Name = "微软雅黑";
        descriptionParagraph.Range.Font.Size = 12;
        descriptionParagraph.SpaceAfter = 15;
        
        // 添加产品特性列表
        AddProductFeatures(document, product.Features);
        
        // 添加价格信息
        var priceParagraph = document.AddParagraph(document.Content.End - 1, 
            $"价格: ¥{product.Price.ToString("N2")}");
        priceParagraph.Range.Font.Name = "微软雅黑";
        priceParagraph.Range.Font.Size = 16;
        priceParagraph.Range.Font.Bold = 1;
        priceParagraph.Range.Font.Color = WdColor.wdColorRed;
        priceParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
    }
    
    /// <summary>
    /// 添加产品特性列表
    /// </summary>
    /// <param name="document">文档对象</param>
    /// <param name="features">特性列表</param>
    private void AddProductFeatures(IWordDocument document, List<string> features)
    {
        // 添加标题
        var featuresTitle = document.AddParagraph(document.Content.End - 1, "产品特性:");
        featuresTitle.Range.Font.Bold = 1;
        featuresTitle.SpaceAfter = 10;
        
        // 添加特性列表
        foreach (var feature in features)
        {
            var featureParagraph = document.AddParagraph(document.Content.End - 1, "• " + feature);
            featureParagraph.Range.Font.Name = "微软雅黑";
            featureParagraph.Range.Font.Size = 11;
            featureParagraph.FirstLineIndent = -20; // 负缩进以对齐项目符号
            featureParagraph.LeftIndent = 20;
            featureParagraph.SpaceAfter = 5;
        }
        
        // 添加空行
        document.AddParagraph(document.Content.End - 1);
    }
}

/// <summary>
/// 产品信息模型
/// </summary>
public class Product
{
    /// <summary>
    /// 产品名称
    /// </summary>
    public string Name { get; set; }
    
    /// <summary>
    /// 产品描述
    /// </summary>
    public string Description { get; set; }
    
    /// <summary>
    /// 产品特性列表
    /// </summary>
    public List<string> Features { get; set; }
    
    /// <summary>
    /// 产品价格
    /// </summary>
    public decimal Price { get; set; }
    
    /// <summary>
    /// 产品图片路径
    /// </summary>
    public string ImagePath { get; set; }
}
```

## 实战案例：创建员工信息表

现在，让我们通过一个完整的示例来综合运用所学知识，创建一个包含多种结构化元素的员工信息表。

```csharp
using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System;

public class EmployeeInfoReportGenerator
{
    /// <summary>
    /// 生成员工信息报告
    /// </summary>
    public void GenerateEmployeeReport()
    {
        try
        {
            // 创建新的Word文档
            using var wordApp = WordFactory.BlankWorkbook();
            var document = wordApp.ActiveDocument;
            wordApp.Visibility = WordAppVisibility.Hidden;
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // 设置文档页面布局
            document.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
            document.PageSetup.TopMargin = 72;    // 1英寸 = 72磅
            document.PageSetup.BottomMargin = 72;
            document.PageSetup.LeftMargin = 72;
            document.PageSetup.RightMargin = 72;

            // 添加标题
            var titleParagraph = document.AddParagraph(0, "员工信息报告");
            titleParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            titleParagraph.Range.Font.Name = "微软雅黑";
            titleParagraph.Range.Font.Size = 20;
            titleParagraph.Range.Font.Bold = 1;

            // 添加空行
            document.AddParagraph(document.Content.End - 1);

            // 添加报告日期
            var dateParagraph = document.AddParagraph(document.Content.End - 1, $"生成日期: {DateTime.Now:yyyy年MM月dd日}");
            dateParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            dateParagraph.Range.Font.Size = 12;

            // 添加空行
            document.AddParagraph(document.Content.End - 1);
            document.AddParagraph(document.Content.End - 1);

            // 创建员工信息表格
            var tableRange = document.Content;
            tableRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            var table = document.Tables.Add(tableRange, 6, 5); // 5行数据+1行标题

            // 设置表格标题行
            string[] headers = { "员工编号", "姓名", "部门", "职位", "入职日期" };
            for (int i = 1; i <= headers.Length; i++)
            {
                var cell = table.Cell(1, i);
                cell.Range.Text = headers[i - 1];
                cell.Range.Font.Bold = 1;
                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }

            // 填充表格数据
            string[,] employeeData = {
                {"E001", "张三", "技术部", "高级软件工程师", "2020-03-15"},
                {"E002", "李四", "市场部", "市场经理", "2019-07-22"},
                {"E003", "王五", "人事部", "人事专员", "2021-01-10"},
                {"E004", "赵六", "财务部", "财务主管", "2018-11-05"},
                {"E005", "钱七", "技术部", "前端开发工程师", "2022-02-28"}
            };

            for (int i = 0; i < employeeData.GetLength(0); i++)
            {
                for (int j = 0; j < employeeData.GetLength(1); j++)
                {
                    var cell = table.Cell(i + 2, j + 1);
                    cell.Range.Text = employeeData[i, j];
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            // 设置表格样式
            table.Borders.Enable = 1;
            table.Borders.LineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.LineWidth = WdLineWidth.wdLineWidth100pt;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100;

            // 调整列宽
            table.Columns[1].PreferredWidth = 15;  // 员工编号列
            table.Columns[2].PreferredWidth = 20;  // 姓名列
            table.Columns[3].PreferredWidth = 20;  // 部门列
            table.Columns[4].PreferredWidth = 25;  // 职位列
            table.Columns[5].PreferredWidth = 20;  // 入职日期列

            // 添加总结段落
            document.AddParagraph(document.Content.End - 1);
            var summaryParagraph = document.AddParagraph(document.Content.End - 1,
                $"本报告共包含 {employeeData.GetLength(0)} 名员工的信息。所有数据均为最新更新。");
            summaryParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            summaryParagraph.FirstLineIndent = 21; // 首行缩进
            summaryParagraph.SpaceBefore = 12;
            summaryParagraph.SpaceAfter = 12;

            // 插入公司Logo（如果存在）
            try
            {
                if (System.IO.File.Exists(@"C:\Images\company_logo.png"))
                {
                    document.AddParagraph(document.Content.End - 1);
                    var logoRange = document.Content;
                    logoRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    var logoShape = document.InlineShapes.AddPicture(@"C:\Images\company_logo.png", false, true);
                    logoShape.Width = 120;
                    logoShape.Height = 60;
                    logoRange.InsertAfter("\n");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"插入Logo时发生错误: {ex.Message}");
            }

            // 保存文档
            string outputPath = $@"C:\Reports\EmployeeReport_{DateTime.Now:yyyyMMdd}.docx";
            document.SaveAs(outputPath, WdSaveFormat.wdFormatXMLDocument);
            document.Close();

            Console.WriteLine($"员工信息报告已生成: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成员工信息报告时发生错误: {ex.Message}");
        }
    }
}

// 使用示例
class Program
{
    static void Main(string[] args)
    {
        var generator = new EmployeeInfoReportGenerator();
        generator.GenerateEmployeeReport();
    }
}
```

## 总结

本文详细介绍了如何使用MudTools.OfficeInterop.Word库操作Word文档中的结构化元素，包括：

1. **段落与节操作**：学习了如何遍历文档中的所有段落，以及如何使用节为文档的不同部分设置不同的页面布局。

2. **表格自动化**：掌握了创建表格、填充数据、设置样式和边框，以及合并拆分单元格等操作。

3. **图片与形状插入**：了解了内嵌形状和浮动形状的区别，以及如何插入图片并设置其属性。

通过实战案例，我们综合运用了这些知识点，创建了一个完整的员工信息报告。这些技能对于开发文档自动化系统、报告生成工具等应用具有重要意义。

在下一篇文章中，我们将学习Word文档的高级格式化技巧，包括样式应用、模板使用、邮件合并等高级功能，敬请期待！