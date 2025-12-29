# 第5章：文本格式化

在Word文档处理中，文本格式化是提升文档可读性和专业性的关键环节。MudTools.OfficeInterop.Word库提供了丰富的文本格式化功能，包括字体、段落、样式等各个方面。本章将详细介绍如何使用这些功能创建格式精美的文档。

## 字体格式设置

字体格式化是文本格式化的基础，包括字体类型、大小、颜色、粗体、斜体等属性。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 获取文档范围
var range = document.Range();

// 添加标题文本
range.Text = "文档标题\n";
```

首先创建文档并添加标题文本。

```csharp
range.Font.Name = "微软雅黑";
range.Font.Size = 18;
range.Font.Bold = 1;
range.Font.Color = WdColor.wdColorBlue;
```

设置字体格式属性：
- Name：字体名称，设置为"微软雅黑"
- Size：字体大小，设置为18磅
- Bold：粗体，1表示开启，0表示关闭
- Color：字体颜色，设置为蓝色

```csharp
// 添加正文内容
var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
contentRange.Text = "这是文档的正文内容，使用标准字体格式。\n";
contentRange.Font.Name = "宋体";
contentRange.Font.Size = 12;
contentRange.Font.Bold = 0;
contentRange.Font.Italic = 0;
```

添加正文内容并设置标准格式：
- 使用"宋体"字体
- 字体大小为12磅
- 非粗体、非斜体

## 段落格式设置

段落格式化涉及对齐方式、缩进、行距、段前段后间距等属性。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加标题
var titleRange = document.Range();
titleRange.Text = "居中标题\n";
titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
titleRange.ParagraphFormat.SpaceAfter = 12;
```

设置居中标题：
- Alignment：段落对齐方式，设置为居中对齐
- SpaceAfter：段后间距，设置为12磅

```csharp
// 添加左对齐段落
var leftPara = document.Range(document.Content.End - 1, document.Content.End - 1);
leftPara.Text = "这是左对齐的段落文本。\n";
leftPara.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

// 添加右对齐段落
var rightPara = document.Range(document.Content.End - 1, document.Content.End - 1);
rightPara.Text = "这是右对齐的段落文本。\n";
rightPara.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

// 添加两端对齐段落
var justifyPara = document.Range(document.Content.End - 1, document.Content.End - 1);
justifyPara.Text = "这是两端对齐的段落文本，文本会自动调整以填满整行。\n";
justifyPara.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
```

设置不同对齐方式的段落：
- 左对齐：wdAlignParagraphLeft
- 右对齐：wdAlignParagraphRight
- 两端对齐：wdAlignParagraphJustify

## 样式应用

样式是格式化的快捷方式，可以一次性应用一组预定义的格式设置。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 应用内置样式
var heading1 = document.Range();
heading1.Text = "标题 1\n";
heading1.Style = "标题 1";
```

应用Word内置的"标题 1"样式。

```csharp
var heading2 = document.Range(document.Content.End - 1, document.Content.End - 1);
heading2.Text = "标题 2\n";
heading2.Style = "标题 2";

var normalText = document.Range(document.Content.End - 1, document.Content.End - 1);
normalText.Text = "正文文本\n";
normalText.Style = "正文";
```

应用其他内置样式。

```csharp
// 创建自定义样式
var customStyle = document.Styles.Add("我的自定义样式");
customStyle.Font.Name = "楷体";
customStyle.Font.Size = 14;
customStyle.Font.Color = WdColor.wdColorRed;
customStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

创建自定义样式并设置其属性：
- 字体：楷体
- 字号：14磅
- 颜色：红色
- 对齐：居中

```csharp
var customRange = document.Range(document.Content.End - 1, document.Content.End - 1);
customRange.Text = "使用自定义样式的文本\n";
customRange.Style = "我的自定义样式";
```

应用自定义样式到文本。

## 列表和编号

列表和编号可以有效组织文档内容，提升可读性。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 创建项目符号列表
var bulletList = document.Range();
bulletList.Text = "项目 1\n项目 2\n项目 3\n";
bulletList.ListFormat.ApplyBulletDefault();
```

创建项目符号列表：
- 添加列表项文本
- 使用ApplyBulletDefault()应用默认的项目符号格式

```csharp
// 创建编号列表
var numberedList = document.Range(document.Content.End - 1, document.Content.End - 1);
numberedList.Text = "第一项\n第二项\n第三项\n";
numberedList.ListFormat.ApplyNumberDefault();
```

创建编号列表：
- 添加列表项文本
- 使用ApplyNumberDefault()应用默认的编号格式

```csharp
// 创建多级列表
var multiLevelList = document.Range(document.Content.End - 1, document.Content.End - 1);
multiLevelList.Text = "主要项目\n子项目 1\n子项目 2\n另一个主要项目\n其子项目\n";
multiLevelList.ListFormat.ApplyOutlineNumberDefault();
```

创建多级列表：
- 添加具有层次结构的文本
- 使用ApplyOutlineNumberDefault()应用大纲编号格式

## 边框和底纹

边框和底纹可以突出显示重要内容或分隔不同部分。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加文本
var range = document.Range();
range.Text = "带边框和底纹的文本\n";

// 设置边框
range.Borders.Enable = 1;
range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
```

设置文本边框：
- Enable = 1：启用边框
- 分别设置上、左、下、右边框为单线样式

```csharp
// 设置底纹
range.Shading.BackgroundPatternColor = WdColor.wdColorLightYellow;
```

设置底纹背景色为浅黄色。

## 制表符设置

制表符用于对齐文本，创建表格效果。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 设置制表符
var range = document.Range();
range.Text = "姓名\t年龄\t职业\n张三\t25\t工程师\n李四\t30\t设计师\n";
```

使用制表符(\t)分隔列数据。

```csharp
// 在特定位置添加制表符
range.ParagraphFormat.TabStops.Add(100, WdTabAlignment.wdAlignTabLeft);
range.ParagraphFormat.TabStops.Add(200, WdTabAlignment.wdAlignTabLeft);
range.ParagraphFormat.TabStops.Add(300, WdTabAlignment.wdAlignTabLeft);
```

添加制表符位置：
- 在100磅、200磅、300磅位置添加左对齐制表符

## 实际应用示例

以下示例演示了如何综合运用各种格式化功能创建专业文档：

```csharp
using MudTools.OfficeInterop;
using System;

class DocumentFormattingDemo
{
    public static void CreateFormattedDocument()
    {
        using var app = WordFactory.BlankDocument();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            
            // 设置文档属性
            document.Title = "格式化文档示例";
            
            // 添加标题
            var title = document.Range();
            title.Text = "公司年度报告\n";
            title.Font.Name = "微软雅黑";
            title.Font.Size = 24;
            title.Font.Bold = 1;
            title.Font.Color = WdColor.wdColorDarkBlue;
            title.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            title.ParagraphFormat.SpaceAfter = 24;
```

设置文档标题格式：
- 字体：微软雅黑，24磅，粗体，深蓝色
- 对齐：居中对齐
- 段后间距：24磅

```csharp
            // 添加副标题
            var subtitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            subtitle.Text = "2025财年总结\n\n";
            subtitle.Font.Name = "微软雅黑";
            subtitle.Font.Size = 16;
            subtitle.Font.Bold = 1;
            subtitle.Font.Color = WdColor.wdColorBlue;
            subtitle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            subtitle.ParagraphFormat.SpaceAfter = 18;
```

设置副标题格式：
- 字体：微软雅黑，16磅，粗体，蓝色
- 对齐：居中对齐
- 段后间距：18磅

```csharp
            // 添加章节标题
            var sectionTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            sectionTitle.Text = "财务概览\n";
            sectionTitle.Font.Name = "微软雅黑";
            sectionTitle.Font.Size = 14;
            sectionTitle.Font.Bold = 1;
            sectionTitle.ParagraphFormat.SpaceAfter = 12;
```

设置章节标题格式：
- 字体：微软雅黑，14磅，粗体
- 段后间距：12磅

```csharp
            // 添加正文内容
            var content = document.Range(document.Content.End - 1, document.Content.End - 1);
            content.Text = "本年度公司实现了显著的财务增长，总收入达到1.2亿元，同比增长15%。净利润为3000万元，同比增长20%。\n\n";
            content.Font.Name = "宋体";
            content.Font.Size = 12;
            content.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            content.ParagraphFormat.FirstLineIndent = 21; // 首行缩进
```

设置正文内容格式：
- 字体：宋体，12磅
- 对齐：两端对齐
- 首行缩进：21磅（约等于中文两个字符的宽度）

```csharp
            // 添加要点列表
            var points = document.Range(document.Content.End - 1, document.Content.End - 1);
            points.Text = "收入增长主要来源：\n• 产品线扩展\n• 市场份额提升\n• 新客户获取\n";
            points.Font.Name = "宋体";
            points.Font.Size = 12;
            
            // 添加表格数据
            var tableSection = document.Range(document.Content.End - 1, document.Content.End - 1);
            tableSection.Text = "\n关键财务指标：\n";
            tableSection.Font.Name = "微软雅黑";
            tableSection.Font.Size = 13;
            tableSection.Font.Bold = 1;
```

添加列表和表格标题，并设置相应格式。

```csharp
            // 创建表格
            var tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            var table = document.Tables.Add(tableRange, 4, 3);
            table.Cell(1, 1).Range.Text = "指标";
            table.Cell(1, 2).Range.Text = "2024年";
            table.Cell(1, 3).Range.Text = "2025年";
            table.Cell(2, 1).Range.Text = "总收入(万元)";
            table.Cell(2, 2).Range.Text = "10,000";
            table.Cell(2, 3).Range.Text = "12,000";
            table.Cell(3, 1).Range.Text = "净利润(万元)";
            table.Cell(3, 2).Range.Text = "2,500";
            table.Cell(3, 3).Range.Text = "3,000";
            table.Cell(4, 1).Range.Text = "增长率";
            table.Cell(4, 2).Range.Text = "-";
            table.Cell(4, 3).Range.Text = "20%";
```

创建并填充表格数据。

```csharp
            // 格式化表格
            table.Borders.Enable = 1;
            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            
            // 保存文档
            document.SaveAs2(@"C:\temp\FormattedDocumentDemo.docx");
            
            Console.WriteLine($"格式化文档已创建: {document.FullName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建文档时出错: {ex.Message}");
        }
    }
}
```

设置表格格式并保存文档。

## 应用场景

1. **企业报告**：创建专业的企业财务报告、年度总结等
2. **学术论文**：格式化论文标题、摘要、正文和参考文献
3. **合同文档**：标准化合同格式，确保法律文档的专业性
4. **营销材料**：制作宣传册、产品介绍等营销文档

## 要点总结

- 字体格式化包括字体类型、大小、颜色、粗体、斜体等基本属性
- 段落格式化涉及对齐方式、缩进、行距等布局属性
- 样式提供了一种快速应用一组格式设置的方法
- 列表和编号有助于组织和结构化文档内容
- 边框和底纹可以突出显示重要信息
- 制表符用于创建对齐的文本布局

掌握文本格式化技能对于创建专业、美观的Word文档至关重要，这些功能使开发者能够自动化生成符合企业或行业标准的文档。