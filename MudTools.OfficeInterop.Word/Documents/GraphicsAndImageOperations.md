# 第7章：图形和图片操作

在Word文档中，图形和图片是增强文档视觉效果、提升表达力的重要元素。MudTools.OfficeInterop.Word库提供了丰富的图形和图片操作功能，包括插入图片、创建形状、使用SmartArt图形等。本章将详细介绍如何使用这些功能为文档添加视觉元素。

## 图形对象管理

Word中的图形对象包括图片、形状、艺术字、SmartArt等。这些对象可以通过InlineShapes和Shapes集合进行管理。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 获取内嵌图形集合
var inlineShapes = document.InlineShapes;
```

InlineShapes属性返回文档中所有内嵌图形的集合，这些图形与文本在同一行。

```csharp
// 获取浮动图形集合
var shapes = document.Shapes;
```

Shapes属性返回文档中所有浮动图形的集合，这些图形可以自由放置在页面上。

```csharp
// 获取图形数量
int inlineShapeCount = inlineShapes.Count;
int shapeCount = shapes.Count;

Console.WriteLine($"内嵌图形数量: {inlineShapeCount}");
Console.WriteLine($"浮动图形数量: {shapeCount}");
```

通过Count属性获取各类图形的数量。

## 图片插入和调整

插入图片是文档处理中最常用的操作之一，可以通过多种方式插入并调整图片。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 在文档末尾插入图片
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var inlineShape = range.InlineShapes.AddPicture(@"C:\images\example.jpg");
```

使用InlineShapes.AddPicture方法插入图片：
- 参数为图片文件的完整路径
- 图片以内嵌方式插入，与文本在同一行

```csharp
// 设置图片尺寸
inlineShape.Width = 300;
inlineShape.Height = 200;
```

设置图片的宽度为300磅，高度为200磅。

```csharp
// 保持图片纵横比
inlineShape.LockAspectRatio = MsoTriState.msoTrue;
```

LockAspectRatio属性控制是否保持图片的原始纵横比，避免图片变形。

```csharp
// 设置图片替代文本
inlineShape.AlternativeText = "示例图片";
```

AlternativeText属性设置图片的替代文本，有助于无障碍访问。

```csharp
// 调整图片亮度和对比度
inlineShape.PictureFormat.Brightness = 0.1f; // 亮度调整
inlineShape.PictureFormat.Contrast = 0.2f;   // 对比度调整
```

通过PictureFormat属性调整图片的亮度和对比度：
- Brightness：亮度值，范围从-1.0到1.0
- Contrast：对比度值，范围从-1.0到1.0

```csharp
// 设置图片环绕方式
// 将内嵌图片转换为浮动图片以支持环绕
var shape = inlineShape.ConvertToShape();
shape.WrapFormat.Type = WdWrapType.wdWrapSquare;
```

转换内嵌图片为浮动图片并设置环绕方式：
- ConvertToShape()：将内嵌图形转换为浮动图形
- WrapFormat.Type：设置环绕类型为方形环绕

## 形状操作

Word支持多种形状，包括基本形状、箭头、流程图元素等。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加矩形形状
var shape1 = document.Shapes.AddShape(
    MsoAutoShapeType.msoShapeRectangle,
    100, 100, 200, 100);
```

使用Shapes.AddShape方法添加矩形：
- 第一个参数：形状类型为矩形
- 第二个参数：左侧位置（100磅）
- 第三个参数：顶部位置（100磅）
- 第四个参数：宽度（200磅）
- 第五个参数：高度（100磅）

```csharp
// 设置形状文本
shape1.TextFrame.TextRange.Text = "矩形形状";
```

通过TextFrame.TextRange.Text属性设置形状内的文本内容。

```csharp
// 设置形状填充
shape1.Fill.ForeColor.RGB = (int)WdColor.wdColorBlue;
```

设置形状填充颜色为蓝色。

```csharp
// 设置形状边框
shape1.Line.ForeColor.RGB = (int)WdColor.wdColorBlack;
shape1.Line.Weight = 2;
```

设置形状边框：
- 边框颜色为黑色
- 边框粗细为2磅

```csharp
// 添加圆形形状
var shape2 = document.Shapes.AddShape(
    MsoAutoShapeType.msoShapeOval,
    150, 250, 150, 150);

shape2.TextFrame.TextRange.Text = "圆形";
shape2.Fill.ForeColor.RGB = (int)WdColor.wdColorRed;
```

添加圆形形状并设置文本和填充颜色。

```csharp
// 添加箭头形状
var shape3 = document.Shapes.AddShape(
    MsoAutoShapeType.msoShapeRightArrow,
    100, 450, 200, 50);

shape3.TextFrame.TextRange.Text = "箭头";
shape3.Fill.ForeColor.RGB = (int)WdColor.wdColorGreen;
```

添加右箭头形状并设置文本和填充颜色。

## SmartArt图形

SmartArt是Word中用于创建专业图表的工具，可以快速创建各种类型的图形。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加SmartArt图形
var range = document.Range(document.Content.End - 1, document.Content.End - 1);
var smartArtShape = document.Shapes.AddSmartArt(
    MsoSmartArtDefaultConstants.msoSmartArtDefaultCycle,
    100, 100, 300, 300);
```

使用Shapes.AddSmartArt方法添加SmartArt图形：
- 第一个参数：SmartArt类型为循环图
- 后续参数：位置和尺寸信息

```csharp
// 获取SmartArt对象
var smartArt = smartArtShape.SmartArt;
```

通过Shape的SmartArt属性获取SmartArt对象。

```csharp
// 添加节点文本
if (smartArt.AllNodes.Count > 0)
{
    smartArt.AllNodes[1].TextFrame.TextRange.Text = "步骤1";
}

if (smartArt.AllNodes.Count > 1)
{
    smartArt.AllNodes[2].TextFrame.TextRange.Text = "步骤2";
}

if (smartArt.AllNodes.Count > 2)
{
    smartArt.AllNodes[3].TextFrame.TextRange.Text = "步骤3";
}
```

为SmartArt的各个节点添加文本内容。

```csharp
// 设置SmartArt颜色样式
smartArt.Color = smartArt.Parent.SmartArtColors[2];
```

设置SmartArt的颜色样式。

```csharp
// 设置SmartArt布局样式
smartArt.Layout = smartArt.Parent.SmartArtLayouts[3];
```

设置SmartArt的布局样式。

## 图形效果设置

可以为图形添加各种视觉效果，如阴影、发光、反射等。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加形状
var shape = document.Shapes.AddShape(
    MsoAutoShapeType.msoShapeRoundedRectangle,
    100, 100, 200, 100);

shape.TextFrame.TextRange.Text = "带效果的形状";
```

添加圆角矩形并设置文本。

```csharp
// 设置阴影效果
shape.Shadow.Visible = MsoTriState.msoTrue;
shape.Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow;
shape.Shadow.Blur = 5;
shape.Shadow.OffsetX = 3;
shape.Shadow.OffsetY = 3;
shape.Shadow.ForeColor.RGB = (int)WdColor.wdColorGray50;
```

设置阴影效果：
- Visible：设置阴影可见
- Style：设置阴影样式为外阴影
- Blur：设置阴影模糊度为5磅
- OffsetX/Y：设置阴影偏移量
- ForeColor：设置阴影颜色为50%灰色

```csharp
// 设置发光效果
shape.Glow.Radius = 5;
shape.Glow.Color.RGB = (int)WdColor.wdColorBlue;
```

设置发光效果：
- Radius：发光半径为5磅
- Color：发光颜色为蓝色

```csharp
// 设置柔化边缘效果
shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType6;
shape.SoftEdge.Radius = 5;
```

设置柔化边缘效果。

```csharp
// 设置三维格式
shape.ThreeD.Visible = MsoTriState.msoTrue;
shape.ThreeD.BevelTopType = MsoBevelType.msoBevelCircle;
shape.ThreeD.BevelTopInset = 3;
shape.ThreeD.BevelTopDepth = 2;
```

设置三维效果：
- Visible：启用三维效果
- BevelTopType：设置顶部斜面类型为圆形
- BevelTopInset：设置顶部斜面内缩为3磅
- BevelTopDepth：设置顶部斜面深度为2磅

## 实际应用示例

以下示例演示了如何创建一个包含多种图形元素的文档：

```csharp
using MudTools.OfficeInterop;
using System;

class GraphicsDemo
{
    public static void CreateGraphicsDocument()
    {
        using var app = WordFactory.BlankDocument();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            
            // 添加标题
            var title = document.Range();
            title.Text = "图形和图片操作示例\n";
            title.Font.Name = "微软雅黑";
            title.Font.Size = 18;
            title.Font.Bold = 1;
            title.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            title.ParagraphFormat.SpaceAfter = 24;
```

设置文档标题格式。

```csharp
            // 添加说明文字
            var description = document.Range(document.Content.End - 1, document.Content.End - 1);
            description.Text = "本文档展示了如何在Word中使用MudTools.OfficeInterop.Word库操作图形和图片\n\n";
            description.Font.Name = "宋体";
            description.Font.Size = 12;
            
            // 添加图片部分标题
            var imageTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            imageTitle.Text = "1. 图片操作示例\n";
            imageTitle.Font.Bold = 1;
            imageTitle.Font.Size = 14;
            imageTitle.ParagraphFormat.SpaceAfter = 12;
```

添加章节标题。

```csharp
            // 添加图片说明
            var imageDescription = document.Range(document.Content.End - 1, document.Content.End - 1);
            imageDescription.Text = "以下是一张示例图片，展示了图片插入和调整功能：\n";
            imageDescription.Font.Name = "宋体";
            imageDescription.Font.Size = 12;
            
            // 插入图片（如果图片存在）
            try
            {
                var imageRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                var inlineShape = imageRange.InlineShapes.AddPicture(@"C:\Windows\Web\Wallpaper\Windows\img0.jpg");
```

尝试插入系统自带的示例图片。

```csharp
                // 调整图片大小
                inlineShape.Width = 400;
                inlineShape.Height = 300;
                inlineShape.LockAspectRatio = MsoTriState.msoTrue;
```

调整图片尺寸并保持纵横比。

```csharp
                // 添加图片说明
                var imageCaption = document.Range(document.Content.End - 1, document.Content.End - 1);
                imageCaption.Text = "\n图1: 示例图片\n\n";
                imageCaption.Font.Italic = 1;
                imageCaption.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            catch (Exception ex)
            {
                // 如果图片不存在，添加提示文字
                var noImageText = document.Range(document.Content.End - 1, document.Content.End - 1);
                noImageText.Text = "[图片未找到]\n\n";
                noImageText.Font.Color = WdColor.wdColorRed;
            }
```

处理图片不存在的情况。

```csharp
            // 添加形状部分标题
            var shapeTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            shapeTitle.Text = "2. 形状操作示例\n";
            shapeTitle.Font.Bold = 1;
            shapeTitle.Font.Size = 14;
            shapeTitle.ParagraphFormat.SpaceAfter = 12;
            
            // 添加形状说明
            var shapeDescription = document.Range(document.Content.End - 1, document.Content.End - 1);
            shapeDescription.Text = "以下是一些基本形状示例：\n";
            shapeDescription.Font.Name = "宋体";
            shapeDescription.Font.Size = 12;
            
            // 添加各种形状
            // 矩形
            var rectangle = document.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                100, 100, 150, 75);
            rectangle.TextFrame.TextRange.Text = "矩形";
            rectangle.Fill.ForeColor.RGB = (int)WdColor.wdColorLightBlue;
```

添加矩形形状并设置样式。

```csharp
            // 圆形
            var circle = document.Shapes.AddShape(
                MsoAutoShapeType.msoShapeOval,
                300, 100, 100, 100);
            circle.TextFrame.TextRange.Text = "圆形";
            circle.Fill.ForeColor.RGB = (int)WdColor.wdColorLightGreen;
            
            // 三角形
            var triangle = document.Shapes.AddShape(
                MsoAutoShapeType.msoShapeIsoscelesTriangle,
                450, 100, 100, 100);
            triangle.TextFrame.TextRange.Text = "三角形";
            triangle.Fill.ForeColor.RGB = (int)WdColor.wdColorLightYellow;
```

添加圆形和三角形形状。

```csharp
            // 添加SmartArt部分标题
            var smartArtTitle = document.Range(document.Content.End - 1, document.Content.End - 1);
            smartArtTitle.Text = "\n\n3. SmartArt图形示例\n";
            smartArtTitle.Font.Bold = 1;
            smartArtTitle.Font.Size = 14;
            smartArtTitle.ParagraphFormat.SpaceAfter = 12;
            
            // 添加SmartArt说明
            var smartArtDescription = document.Range(document.Content.End - 1, document.Content.End - 1);
            smartArtDescription.Text = "以下是一个SmartArt图形示例：\n";
            smartArtDescription.Font.Name = "宋体";
            smartArtDescription.Font.Size = 12;
            
            // 添加SmartArt图形
            var smartArtRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            var smartArtShape = document.Shapes.AddSmartArt(
                MsoSmartArtDefaultConstants.msoSmartArtDefaultList,
                Left: 100,
                Top: 100,
                Width: 400,
                Height: 300);
```

添加SmartArt列表图形。

```csharp
            // 获取SmartArt对象并设置文本
            var smartArt = smartArtShape.SmartArt;
            if (smartArt.AllNodes.Count >= 3)
            {
                smartArt.AllNodes[1].TextFrame.TextRange.Text = "项目1: 需求分析";
                smartArt.AllNodes[2].TextFrame.TextRange.Text = "项目2: 系统设计";
                smartArt.AllNodes[3].TextFrame.TextRange.Text = "项目3: 编码实现";
            }
            
            // 保存文档
            document.SaveAs2(@"C:\temp\GraphicsDemo.docx");
            
            Console.WriteLine($"图形文档已创建: {document.FullName}");
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

1. **演示文档**：为演示文稿添加图表、图片等视觉元素
2. **产品手册**：在产品说明中插入产品图片和示意图
3. **教学材料**：通过图形和图片增强教学内容的可视化效果
4. **报告文档**：使用图表展示数据分析结果

## 要点总结

- 图形对象包括图片、形状、SmartArt等多种类型
- 可以通过InlineShapes和Shapes集合管理图形对象
- 图片插入支持多种格式，并可调整尺寸、亮度、对比度等属性
- 形状操作支持多种基本形状，并可设置填充、边框等样式
- SmartArt图形提供了专业的图表创建功能
- 图形效果设置包括阴影、发光、反射等多种视觉效果

掌握图形和图片操作技能对于创建视觉吸引力强的Word文档至关重要，这些功能使开发者能够自动化生成包含丰富视觉元素的专业文档。