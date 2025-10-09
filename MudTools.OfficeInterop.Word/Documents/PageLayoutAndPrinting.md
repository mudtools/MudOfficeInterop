# 第8章：页面布局和打印

页面布局和打印设置是创建专业文档的重要环节。MudTools.OfficeInterop.Word库提供了完整的页面设置和打印控制功能，包括页面尺寸、边距、页眉页脚、分节符等。本章将详细介绍如何使用这些功能创建符合要求的文档格式。

## 页面设置 (IWordPageSetup)

页面设置是文档布局的基础，包括纸张大小、方向、边距等属性。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 获取页面设置对象
var pageSetup = document.Sections[1].PageSetup;
```

通过Sections[1]获取文档第一节的PageSetup对象，第一节通常包含整个文档的页面设置。

```csharp
// 设置纸张大小
pageSetup.PageWidth = 12240; // A4纸宽度 (单位: 磅/72英寸)
pageSetup.PageHeight = 15840; // A4纸高度
```

直接设置页面宽度和高度（以磅为单位）：
- PageWidth：页面宽度，12240磅约等于21厘米（A4宽度）
- PageHeight：页面高度，15840磅约等于29.7厘米（A4高度）

```csharp
// 或者使用预定义的纸张大小
pageSetup.PageSize = WdPaperSize.wdPaperA4;
```

使用预定义的纸张大小常量wdPaperA4设置A4纸张。

```csharp
// 设置页面方向
pageSetup.Orientation = WdOrientation.wdOrientPortrait; // 纵向
// pageSetup.Orientation = WdOrientation.wdOrientLandscape; // 横向
```

设置页面方向：
- wdOrientPortrait：纵向（默认）
- wdOrientLandscape：横向

```csharp
// 设置页边距
pageSetup.TopMargin = 1440;    // 1英寸 = 72磅
pageSetup.BottomMargin = 1440;
pageSetup.LeftMargin = 1800;   // 1.25英寸
pageSetup.RightMargin = 1800;
pageSetup.HeaderDistance = 720; // 页眉距离
pageSetup.FooterDistance = 720; // 页脚距离
```

设置页面边距（以磅为单位）：
- TopMargin/BottomMargin：上下边距各1英寸（72磅）
- LeftMargin/RightMargin：左右边距各1.25英寸（1800磅）
- HeaderDistance/FooterDistance：页眉页脚距离页面边缘的距离

```csharp
// 设置页面垂直对齐方式
pageSetup.VerticalAlignment = WdVerticalAlignment.wdAlignVerticalTop;
```

设置页面内容的垂直对齐方式为顶部对齐。

```csharp
// 设置行号
pageSetup.LineNumbering.Active = 1; // 启用行号
pageSetup.LineNumbering.RestartMode = WdNumberingRule.wdRestartContinuous;
```

设置行号：
- Active = 1：启用行号显示
- RestartMode：设置行号重新开始的规则为连续编号

## 页眉和页脚

页眉和页脚是文档中重复出现的内容，通常包含页码、文档标题、日期等信息。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 获取页眉和页脚范围
var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
```

获取第一节主页面的页眉和页脚范围：
- Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary]：获取主页面的页眉
- Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary]：获取主页面的页脚

```csharp
// 设置页眉内容
headerRange.Text = "文档标题";
headerRange.Font.Name = "微软雅黑";
headerRange.Font.Size = 12;
headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

设置页眉内容和格式：
- Text：设置页眉文本内容
- Font.Name/Size：设置字体和字号
- ParagraphFormat.Alignment：设置段落居中对齐

```csharp
// 设置页脚内容（包含页码）
footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage); // 插入页码
footerRange.Text = " 第 ";
footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
footerRange.Fields.Add(footerRange, WdFieldType.wdFieldNumPages); // 插入总页数
footerRange.Text = " 页";
footerRange.Font.Size = 10;
footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

设置包含页码的页脚：
1. Fields.Add(footerRange, WdFieldType.wdFieldPage)：插入当前页码字段
2. 添加文本" 第 "
3. Collapse(WdCollapseDirection.wdCollapseEnd)：将光标移到末尾
4. Fields.Add(footerRange, WdFieldType.wdFieldNumPages)：插入总页数字段
5. 添加文本" 页"

```csharp
// 设置首页不同
document.Sections[1].PageSetup.DifferentFirstPageHeaderFooter = 1;
```

启用首页不同的页眉页脚设置。

```csharp
// 设置奇偶页不同
document.Sections[1].PageSetup.OddAndEvenPagesHeaderFooter = 1;
```

启用奇偶页不同的页眉页脚设置。

## 分节符和分页符

分节符和分页符用于控制文档的结构和布局。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加内容
var range = document.Range();
range.Text = "第一部分内容\n";
```

添加第一部分内容。

```csharp
// 插入分页符
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.InsertBreak(WdBreakType.wdPageBreak);
```

插入分页符：
- Collapse(WdCollapseDirection.wdCollapseEnd)：将光标移到内容末尾
- InsertBreak(WdBreakType.wdPageBreak)：插入分页符

```csharp
// 添加第二部分内容
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "第二部分内容\n";

// 插入分节符（下一页）
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.InsertBreak(WdBreakType.wdSectionBreakNextPage);
```

插入分节符（下一页）：
- 插入新内容
- 使用InsertBreak(WdBreakType.wdSectionBreakNextPage)插入分节符

```csharp
// 添加第三部分内容
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "第三部分内容\n";

// 为不同节设置不同的页面布局
var section1 = document.Sections[1]; // 第一节
section1.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

var section2 = document.Sections[2]; // 第二节
section2.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

var section3 = document.Sections[3]; // 第三节
section3.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
```

为不同节设置不同的页面方向。

## 打印选项和预览

打印选项控制文档的打印行为，包括打印范围、份数、双面打印等。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 打印预览
app.ActiveWindow.View.Type = WdViewType.wdPrintPreviewView;
```

切换到打印预览视图。

```csharp
// 设置打印选项
var printOptions = app.ActiveDocument.PrintOut(
    Background: false,
    Copies: 2,                      // 打印2份
    PageType: WdPrintOutPages.wdPrintAllPages, // 打印所有页面
    Range: WdPrintOutRange.wdPrintAllDocument, // 打印整个文档
    Item: WdPrintOutItem.wdPrintDocumentContent, // 打印文档内容
    Collate: true                   // 逐份打印
);
```

设置打印选项：
- Background：是否在后台打印
- Copies：打印份数
- PageType：打印页面类型
- Range：打印范围
- Item：打印项目
- Collate：是否逐份打印

```csharp
// 获取打印相关信息
int pagesCount = document.Range().Paragraphs.Count; // 粗略估算页数
Console.WriteLine($"文档大约有 {pagesCount} 页");
```

估算文档页数。

## 实际应用示例

以下示例演示了如何创建一个具有完整页面布局设置的专业文档：

```csharp
using MudTools.OfficeInterop;
using System;

class PageLayoutDemo
{
    public static void CreateProfessionalDocument()
    {
        using var app = WordFactory.BlankWorkbook();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            
            // 设置文档属性
            document.Title = "专业文档示例";
            document.Author = "MudTools.OfficeInterop.Word 用户";
```

设置文档的基本属性。

```csharp
            // 设置第一页的页面布局
            var section1 = document.Sections[1];
            var pageSetup = section1.PageSetup;
            
            // 设置A4纸张
            pageSetup.PageSize = WdPaperSize.wdPaperA4;
            pageSetup.Orientation = WdOrientation.wdOrientPortrait;
```

设置第一节使用A4纵向页面。

```csharp
            // 设置页边距
            pageSetup.TopMargin = 1440;    // 2厘米
            pageSetup.BottomMargin = 1440;
            pageSetup.LeftMargin = 1800;   // 2.5厘米
            pageSetup.RightMargin = 1800;
            pageSetup.HeaderDistance = 720;
            pageSetup.FooterDistance = 720;
```

设置页面边距。

```csharp
            // 设置首页不同
            pageSetup.DifferentFirstPageHeaderFooter = 1;
            
            // 添加封面内容
            var coverRange = document.Range();
            coverRange.Text = "\n\n\n";
            coverRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            
            // 添加标题
            var titleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            titleRange.Text = "公司年度报告\n";
            titleRange.Font.Name = "微软雅黑";
            titleRange.Font.Size = 28;
            titleRange.Font.Bold = 1;
            titleRange.Font.Color = WdColor.wdColorDarkBlue;
            titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            titleRange.ParagraphFormat.SpaceAfter = 24;
```

添加封面标题并设置格式。

```csharp
            // 添加副标题
            var subtitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            subtitleRange.Text = "2025财年总结\n\n\n";
            subtitleRange.Font.Name = "微软雅黑";
            subtitleRange.Font.Size = 18;
            subtitleRange.Font.Color = WdColor.wdColorBlue;
            subtitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 添加公司信息
            var companyRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            companyRange.Text = "某某公司\n";
            companyRange.Font.Name = "宋体";
            companyRange.Font.Size = 14;
            companyRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            var dateRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            dateRange.Text = DateTime.Now.ToString("yyyy年MM月dd日") + "\n";
            dateRange.Font.Name = "宋体";
            dateRange.Font.Size = 12;
            dateRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

添加副标题和公司信息。

```csharp
            // 插入分页符
            var breakRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            breakRange.InsertBreak(WdBreakType.wdPageBreak);
```

插入分页符，开始新页面。

```csharp
            // 设置目录页的页眉页脚
            var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            headerRange.Text = "公司年度报告";
            headerRange.Font.Name = "宋体";
            headerRange.Font.Size = 10;
            headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            footerRange.Text = "第 ";
            footerRange.Collapse(WdCollapseDirection.wdCollapseEnd);
            footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
            footerRange.Text = " 页";
            footerRange.Font.Size = 10;
            footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

设置页眉页脚。

```csharp
            // 添加目录标题
            var tocTitleRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            tocTitleRange.Text = "目录\n";
            tocTitleRange.Font.Name = "微软雅黑";
            tocTitleRange.Font.Size = 16;
            tocTitleRange.Font.Bold = 1;
            tocTitleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            tocTitleRange.ParagraphFormat.SpaceAfter = 24;
            
            // 插入分页符
            var breakRange2 = document.Range(document.Content.End - 1, document.Content.End - 1);
            breakRange2.InsertBreak(WdBreakType.wdPageBreak);
```

添加目录和分页符。

```csharp
            // 添加正文内容
            var contentTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            contentTitle.Text = "第一章：公司概况\n";
            contentTitle.Font.Name = "微软雅黑";
            contentTitle.Font.Size = 14;
            contentTitle.Font.Bold = 1;
            contentTitle.ParagraphFormat.SpaceAfter = 12;
            
            var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            contentRange.Text = "这里是公司概况的内容...\n\n";
            contentRange.Font.Name = "宋体";
            contentRange.Font.Size = 12;
            
            // 添加第二章
            var chapter2Title = document.Range(document.Content.End - 1, document.Content.End - 1);
            chapter2Title.Text = "第二章：财务分析\n";
            chapter2Title.Font.Name = "微软雅黑";
            chapter2Title.Font.Size = 14;
            chapter2Title.Font.Bold = 1;
            chapter2Title.ParagraphFormat.SpaceAfter = 12;
            
            var chapter2Range = document.Range(document.Content.End - 1, document.Content.End - 1);
            chapter2Range.Text = "这里是财务分析的内容...\n\n";
            chapter2Range.Font.Name = "宋体";
            chapter2Range.Font.Size = 12;
```

添加正文内容。

```csharp
            // 插入分节符（下一页）
            var sectionBreakRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            sectionBreakRange.InsertBreak(WdBreakType.wdSectionBreakNextPage);
            
            // 为新节设置横向页面
            var section2 = document.Sections[2];
            section2.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            
            // 添加横向页面内容
            var landscapeTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            landscapeTitle.Text = "财务数据图表\n";
            landscapeTitle.Font.Name = "微软雅黑";
            landscapeTitle.Font.Size = 14;
            landscapeTitle.Font.Bold = 1;
            landscapeTitle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            landscapeTitle.ParagraphFormat.SpaceAfter = 12;
            
            // 保存文档
            document.SaveAs2(@"C:\temp\PageLayoutDemo.docx");
            
            Console.WriteLine($"专业文档已创建: {document.FullName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建文档时出错: {ex.Message}");
        }
    }
}
```

保存文档并输出结果。

## 应用场景

1. **企业报告**：创建具有专业外观的年度报告、财务报表等
2. **学术论文**：设置符合学术规范的论文格式，包括页眉页脚、页码等
3. **合同文档**：为法律文档设置标准的页面布局和页码系统
4. **手册指南**：为用户手册设置目录、章节分隔和专业排版

## 要点总结

- 页面设置是文档布局的基础，包括纸张大小、方向、边距等属性
- 页眉和页脚可以添加重复出现的内容，如页码、文档标题等
- 分节符和分页符用于控制文档结构和不同部分的布局
- 打印选项控制文档的打印行为，包括范围、份数、双面打印等
- 通过合理的页面布局设置可以创建专业、规范的文档

掌握页面布局和打印设置技能对于创建符合要求的Word文档至关重要，这些功能使开发者能够自动化生成满足各种格式要求的专业文档。