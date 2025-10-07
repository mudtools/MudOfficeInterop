# 第6章：表格操作

表格是Word文档中组织和展示数据的重要工具。MudTools.OfficeInterop.Word库提供了完整的表格操作功能，包括创建、格式化、数据处理等。本章将详细介绍如何使用IWordTable和IWordTables接口操作Word表格。

## IWordTable和IWordTables接口

IWordTable接口代表单个表格，而IWordTables接口代表文档中所有表格的集合。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 获取表格集合
var tables = document.Tables;
```

通过Document对象的Tables属性获取表格集合。

```csharp
// 获取表格数量
int tableCount = tables.Count;
```

Count属性返回文档中表格的总数。

```csharp
// 访问特定表格（索引从1开始）
if (tableCount > 0)
{
    var firstTable = tables.Item(1);
    // 操作表格
}
```

通过Item方法和索引（从1开始）访问特定表格。

## 创建和删除表格

可以通过多种方式创建表格，也可以根据需要删除表格。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 方法1：在文档末尾添加表格
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var table1 = document.Tables.Add(range, 3, 4); // 3行4列
```

使用Tables.Add方法创建表格：
- 第一个参数：指定表格插入位置的范围
- 第二个参数：表格行数（3行）
- 第三个参数：表格列数（4列）

```csharp
// 方法2：在指定位置添加表格
var range2 = document.Range(0, 0);
var table2 = document.Tables.Add(range2, 2, 3); // 2行3列
```

在文档开头创建2行3列的表格。

```csharp
// 设置表格标题
table1.Title = "示例表格";
table1.Descr = "这是一个示例表格";
```

设置表格的标题和描述信息，有助于无障碍访问。

```csharp
// 删除表格
// table1.Delete(); // 删除整个表格
```

使用Delete方法可以删除整个表格。

## 表格格式化

表格格式化包括边框、底纹、对齐方式、尺寸等设置。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 创建表格
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var table = document.Tables.Add(range, 4, 3);

// 设置表格边框
table.Borders.Enable = 1;
table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
table.Borders[WdBorderType.wdBorderHorizontal].LineStyle = WdLineStyle.wdLineStyleDot;
table.Borders[WdBorderType.wdBorderVertical].LineStyle = WdLineStyle.wdLineStyleDot;
```

设置表格边框：
- Enable = 1：启用边框
- 分别设置上、左、下、右边框为单线样式
- 设置水平和垂直内部边框为点线样式

```csharp
// 设置表格对齐方式
table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
```

设置表格行的对齐方式为居中。

```csharp
// 设置表格宽度
table.AllowAutoFit = false;
table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
table.PreferredWidth = 100;
```

设置表格宽度属性：
- AllowAutoFit = false：禁用自动调整
- PreferredWidthType：设置宽度类型为百分比
- PreferredWidth：设置宽度为100%（页面宽度）

```csharp
// 设置列宽
table.Columns[1].Width = 100;
table.Columns[2].Width = 150;
table.Columns[3].Width = 200;
```

分别设置各列的宽度。

## 单元格操作

单元格是表格的基本组成单位，可以对单元格进行各种操作。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 创建表格
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var table = document.Tables.Add(range, 3, 3);

// 访问单元格
var cell = table.Cell(1, 1); // 第一行第一列（索引从1开始）
var cellRange = cell.Range;
```

访问特定单元格：
- 使用Cell方法，参数为行号和列号（都从1开始）
- 通过Range属性获取单元格的内容范围

```csharp
// 设置单元格文本
cellRange.Text = "单元格内容";
```

设置单元格的文本内容。

```csharp
// 设置单元格格式
cellRange.Font.Bold = 1;
cellRange.Font.Size = 12;
cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

设置单元格文本格式：
- 粗体字体
- 12磅字号
- 居中对齐

```csharp
// 设置单元格底纹
cell.Shading.BackgroundPatternColor = WdColor.wdColorLightBlue;
```

设置单元格背景色为浅蓝色。

```csharp
// 合并单元格
// table.Cell(1, 1).Merge(table.Cell(1, 2)); // 合并第一行的前两列

// 拆分单元格
// table.Cell(1, 1).Split(2, 2); // 将单元格拆分为2行2列
```

合并和拆分单元格操作（注释掉避免实际执行）。

## 表格数据处理

可以对表格中的数据进行各种处理操作。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 创建表格并填充数据
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var table = document.Tables.Add(range, 4, 3);

// 填充表头
table.Cell(1, 1).Range.Text = "姓名";
table.Cell(1, 2).Range.Text = "年龄";
table.Cell(1, 3).Range.Text = "职业";
```

填充表头行数据。

```csharp
// 填充数据
string[,] data = {
    {"张三", "25", "工程师"},
    {"李四", "30", "设计师"},
    {"王五", "28", "产品经理"}
};

for (int i = 0; i < data.GetLength(0); i++)
{
    for (int j = 0; j < data.GetLength(1); j++)
    {
        table.Cell(i + 2, j + 1).Range.Text = data[i, j];
    }
}
```

使用嵌套循环填充数据行：
- 外层循环遍历行（i从0到2）
- 内层循环遍历列（j从0到2）
- 将数据填充到对应的单元格中（行号从2开始，跳过表头）

```csharp
// 格式化表头
for (int i = 1; i <= 3; i++)
{
    var headerCell = table.Cell(1, i);
    headerCell.Range.Font.Bold = 1;
    headerCell.Range.Font.Color = WdColor.wdColorWhite;
    headerCell.Shading.BackgroundPatternColor = WdColor.wdColorDarkBlue;
    headerCell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
}
```

格式化表头：
- 设置粗体、白色字体
- 设置深蓝色背景
- 设置居中对齐

```csharp
// 格式化数据行
for (int row = 2; row <= 4; row++)
{
    for (int col = 1; col <= 3; col++)
    {
        var cell = table.Cell(row, col);
        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
    }
}
```

格式化数据行内容居中对齐。

## 高级表格操作

表格还支持一些高级操作，如排序、公式计算等。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 创建带数据的表格
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var table = document.Tables.Add(range, 5, 3);

// 填充数据
table.Cell(1, 1).Range.Text = "产品";
table.Cell(1, 2).Range.Text = "销量";
table.Cell(1, 3).Range.Text = "价格";

table.Cell(2, 1).Range.Text = "产品A";
table.Cell(2, 2).Range.Text = "100";
table.Cell(2, 3).Range.Text = "50";

table.Cell(3, 1).Range.Text = "产品B";
table.Cell(3, 2).Range.Text = "200";
table.Cell(3, 3).Range.Text = "30";

table.Cell(4, 1).Range.Text = "产品C";
table.Cell(4, 2).Range.Text = "150";
table.Cell(4, 3).Range.Text = "40";
```

创建并填充示例数据。

```csharp
// 格式化表头
for (int i = 1; i <= 3; i++)
{
    var cell = table.Cell(1, i);
    cell.Range.Font.Bold = 1;
    cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
}
```

格式化表头行。

```csharp
// 添加总计行
table.Cell(5, 1).Range.Text = "总计";
table.Cell(5, 2).Range.Text = "=SUM(ABOVE)"; // 使用公式计算总销量
table.Cell(5, 3).Range.Text = "平均价格";
```

添加总计行并使用公式计算总销量：
- "=SUM(ABOVE)"：计算上方所有数值的总和

```csharp
// 更新表格中的字段（公式）
table.Range.Fields.Update();
```

更新表格中的字段，使公式生效。

## 实际应用示例

以下示例演示了如何创建一个完整的数据报表：

```csharp
using MudTools.OfficeInterop;
using System;

class TableReportDemo
{
    public static void CreateTableReport()
    {
        using var app = WordFactory.BlankWorkbook();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            
            // 添加标题
            var title = document.Range();
            title.Text = "销售数据报表\n";
            title.Font.Name = "微软雅黑";
            title.Font.Size = 18;
            title.Font.Bold = 1;
            title.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            title.ParagraphFormat.SpaceAfter = 24;
```

设置报表标题格式。

```csharp
            // 添加报表说明
            var description = document.Range(document.Content.End - 1, document.Content.End - 1);
            description.Text = "本报表展示了2025年各季度销售数据\n\n";
            description.Font.Name = "宋体";
            description.Font.Size = 12;
            description.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

添加报表说明文字。

```csharp
            // 创建销售数据表格
            var tableRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            var table = document.Tables.Add(tableRange, 6, 5);
            
            // 设置表格标题
            table.Title = "季度销售数据";
            table.Descr = "2025年各季度销售数据表";
```

创建表格并设置标题和描述。

```csharp
            // 填充表头
            string[] headers = { "季度", "产品A", "产品B", "产品C", "总计" };
            for (int i = 0; i < headers.Length; i++)
            {
                table.Cell(1, i + 1).Range.Text = headers[i];
            }
```

使用循环填充表头数据。

```csharp
            // 填充数据
            string[,] salesData = {
                {"Q1", "1000", "1500", "2000", "4500"},
                {"Q2", "1200", "1800", "2200", "5200"},
                {"Q3", "1100", "1600", "2100", "4800"},
                {"Q4", "1300", "1900", "2300", "5500"}
            };

            for (int i = 0; i < salesData.GetLength(0); i++)
            {
                for (int j = 0; j < salesData.GetLength(1); j++)
                {
                    table.Cell(i + 2, j + 1).Range.Text = salesData[i, j];
                }
            }
```

填充各季度销售数据。

```csharp
            // 格式化表格
            // 表格边框
            table.Borders.Enable = 1;
            table.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
```

设置表格边框样式。

```csharp
            // 表头格式
            for (int i = 1; i <= 5; i++)
            {
                var cell = table.Cell(1, i);
                cell.Range.Font.Bold = 1;
                cell.Range.Font.Color = WdColor.wdColorWhite;
                cell.Shading.BackgroundPatternColor = WdColor.wdColorDarkBlue;
                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
```

格式化表头行。

```csharp
            // 数据行格式
            for (int row = 2; row <= 6; row++)
            {
                for (int col = 1; col <= 5; col++)
                {
                    var cell = table.Cell(row, col);
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    
                    // 交替行颜色
                    if (row % 2 == 0)
                    {
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray10;
                    }
                }
            }
```

格式化数据行并设置交替行颜色。

```csharp
            // 设置列宽
            table.AllowAutoFit = false;
            table.Columns[1].Width = 80;   // 季度列
            table.Columns[2].Width = 80;   // 产品A列
            table.Columns[3].Width = 80;   // 产品B列
            table.Columns[4].Width = 80;   // 产品C列
            table.Columns[5].Width = 80;   // 总计列
            
            // 更新公式字段
            table.Range.Fields.Update();
            
            // 保存文档
            document.SaveAs2(@"C:\temp\TableReportDemo.docx");
            
            Console.WriteLine($"表格报表已创建: {document.FullName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建报表时出错: {ex.Message}");
        }
    }
}
```

更新公式并保存文档。

## 应用场景

1. **数据分析报告**：展示统计数据和分析结果
2. **财务报表**：创建资产负债表、损益表等财务文档
3. **产品目录**：组织产品信息和规格参数
4. **时间安排表**：展示项目计划和时间安排

## 要点总结

- IWordTable和IWordTables接口提供了完整的表格操作功能
- 可以通过多种方式创建表格并设置其基本属性
- 表格格式化包括边框、底纹、对齐方式等视觉效果
- 单元格操作允许精确控制每个单元格的内容和格式
- 表格数据处理功能支持填充、计算和公式应用
- 高级操作如排序和公式计算使表格更具实用性

掌握表格操作技能对于创建包含结构化数据的Word文档至关重要，这些功能使开发者能够自动化生成各种数据报表和信息展示文档。