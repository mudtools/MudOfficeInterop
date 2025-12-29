# 第3章：文档结构和范围操作

在Word文档处理中，理解文档结构和范围操作是非常重要的。IWordDocument和IWordRange接口提供了丰富的功能来操作文档内容。本章将详细介绍这些接口及其使用方法。

## IWordDocument接口详解

IWordDocument接口代表一个Word文档，提供了访问文档属性和内容的方法。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;
```

首先创建Word应用程序实例并获取活动文档。

```csharp
// 获取文档基本信息
string name = document.Name;
string fullName = document.FullName;
string path = document.Path;
string title = document.Title;
```

这些属性提供了文档的基本信息：
- Name：仅包含文件名，不包含路径
- FullName：包含完整路径的文件名
- Path：文档所在目录路径
- Title：文档标题

```csharp
// 检查文档状态
bool? saved = document.Saved;
bool? routed = document.Routed;
```

状态属性帮助我们了解文档的当前状态：
- Saved：指示文档是否已保存
- Routed：指示文档是否已发送路由

```csharp
// 设置文档属性
document.Title = "新标题";
```

可以修改文档的属性值，如标题。

## 文档基本属性和元数据

文档包含许多重要属性，可以获取和设置文档的元数据：

```csharp
using var app = WordFactory.CreateFrom(@"C:\templates\ReportTemplate.dotx");
var document = app.ActiveDocument;

// 文档基本信息
Console.WriteLine($"文档名称: {document.Name}");
Console.WriteLine($"完整路径: {document.FullName}");
Console.WriteLine($"文档路径: {document.Path}");
Console.WriteLine($"文档标题: {document.Title}");
```

输出文档的基本信息。

```csharp
// 文档状态
Console.WriteLine($"是否已保存: {document.Saved}");
Console.WriteLine($"是否已发送路由: {document.Routed}");
Console.WriteLine($"是否为主控文档: {document.IsMasterDocument}");
```

检查文档的状态信息。

```csharp
// 字体嵌入设置
document.EmbedTrueTypeFonts = true;
document.SaveSubsetFonts = true;
```

设置字体嵌入选项，确保文档在其他计算机上显示一致。

## 文档生命周期管理

正确管理文档的生命周期对于资源管理和数据完整性非常重要：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

try
{
    // 编辑文档内容
    var range = document.Range();
    range.Text = "文档内容";
    
    // 保存文档
    document.SaveAs2(@"C:\temp\example.docx");
    Console.WriteLine("文档已保存");
}
catch (Exception ex)
{
    Console.WriteLine($"保存文档时出错: {ex.Message}");
}
finally
{
    // 关闭文档
    document.Close();
}
```

在这个示例中，我们：
1. 编辑文档内容
2. 保存文档到指定位置
3. 在finally块中确保文档被正确关闭

## IWordRange接口详解

IWordRange接口代表文档中的一个连续区域，是操作文档内容的基础。

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 获取整个文档的范围
var range = document.Range();
```

使用无参数的Range()方法获取整个文档的范围。

```csharp
// 创建指定位置的范围
var specificRange = document.Range(0, 10);
```

通过指定起始位置(0)和结束位置(10)创建特定范围。

```csharp
// 获取范围的副本
var duplicateRange = range.Duplicate;
```

Duplicate属性创建当前范围的副本，可用于独立操作。

## 范围的选择和定义

范围由起始和结束位置定义，可以通过多种方式操作：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 获取整个文档范围
var range = document.Range();
range.Text = "这是第一段文本。\n这是第二段文本。\n这是第三段文本。";
```

首先填充一些示例文本。

```csharp
// 重新定义范围
range.Start = 0;
range.End = 5;
```

通过设置Start和End属性重新定义范围的边界。

```csharp
// 获取范围文本
string text = range.Text;
Console.WriteLine($"范围文本: {text}");
```

获取并输出当前范围内的文本内容。

```csharp
// 移动范围
range.Start = 6;
range.End = 12;
Console.WriteLine($"新范围文本: {range.Text}");
```

移动范围到新的位置并输出文本。

## 文本内容操作

范围提供了丰富的文本操作功能：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;
var range = document.Range();

// 设置文本
range.Text = "Hello World!";
```

直接设置范围的文本内容。

```csharp
// 插入文本
range.InsertBefore("前缀文本 ");
range.InsertAfter(" 后缀文本");
```

使用InsertBefore和InsertAfter方法在范围前后插入文本。

```csharp
// 删除文本
var deleteRange = document.Range(0, 5);
deleteRange.Delete();
```

创建特定范围并删除其中的文本。

```csharp
// 替换文本
range.Text = range.Text.Replace("Hello", "Hi");
```

使用字符串的Replace方法替换文本内容。

## 范围复制和移动

可以复制和移动范围内容：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;

// 添加内容
var range1 = document.Range();
range1.Text = "原始内容\n";

var range2 = document.Range();
range2.Text = "另一部分内容\n";
```

添加两段不同的内容。

```csharp
// 复制内容
var sourceRange = document.Range(0, 4);
var targetRange = document.Range(document.StoryLength, document.StoryLength);
sourceRange.Copy();
targetRange.Paste();
```

复制操作步骤：
1. 选择要复制的源范围
2. 选择目标位置
3. 执行复制操作
4. 在目标位置粘贴

```csharp
// 移动内容
var moveSource = document.Range(5, 9);
var moveTarget = document.Range(document.StoryLength, document.StoryLength);
moveSource.Cut();
moveTarget.Paste();
```

移动操作步骤：
1. 选择要移动的源范围
2. 选择目标位置
3. 执行剪切操作
4. 在目标位置粘贴

## 实际应用示例

以下示例演示了如何综合使用文档和范围操作：

```csharp
using MudTools.OfficeInterop;
using System;

class DocumentProcessor
{
    public static void ProcessDocument()
    {
        using var app = WordFactory.BlankDocument();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            
            // 设置文档属性
            document.Title = "示例文档";
            document.Author = "MudTools.OfficeInterop.Word 用户";
            
            // 添加内容
            var range = document.Range();
            range.Text = "文档标题\n\n";
            range.Font.Bold = 1;
            range.Font.Size = 16;
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            
            // 添加正文
            var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            contentRange.Text = "这是文档的正文内容。\n";
            contentRange.Font.Bold = 0;
            contentRange.Font.Size = 12;
            
            // 添加列表
            var listRange = document.Range(document.Content.End - 1, document.Content.End - 1);
            listRange.Text = "项目1\n项目2\n项目3\n";
            listRange.ListFormat.ApplyBulletDefault();
            
            // 保存文档
            document.SaveAs2(@"C:\temp\StructuredDocument.docx");
            
            Console.WriteLine($"文档已创建: {document.FullName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"处理文档时出错: {ex.Message}");
        }
    }
}
```

让我们逐步分析这个示例：

```csharp
// 设置文档属性
document.Title = "示例文档";
document.Author = "MudTools.OfficeInterop.Word 用户";
```

设置文档的元数据属性。

```csharp
// 添加内容
var range = document.Range();
range.Text = "文档标题\n\n";
range.Font.Bold = 1;
range.Font.Size = 16;
range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

添加标题并设置格式：
1. 设置文本内容
2. 设置粗体（Bold = 1表示开启）
3. 设置字体大小为16磅
4. 设置段落居中对齐

```csharp
// 添加正文
var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
contentRange.Text = "这是文档的正文内容。\n";
contentRange.Font.Bold = 0;
contentRange.Font.Size = 12;
```

在文档末尾添加正文内容并设置格式。

```csharp
// 添加列表
var listRange = document.Range(document.Content.End - 1, document.Content.End - 1);
listRange.Text = "项目1\n项目2\n项目3\n";
listRange.ListFormat.ApplyBulletDefault();
```

添加项目符号列表。

## 应用场景

1. **内容提取工具**：通过范围操作提取文档特定部分的内容
2. **格式化工具**：使用范围功能批量格式化文档内容
3. **文档分析器**：分析文档结构和内容分布
4. **模板引擎**：基于模板动态生成文档内容

## 要点总结

- IWordDocument接口提供了对Word文档的完整访问能力
- 文档属性包含丰富的元数据信息
- 正确管理文档生命周期对于资源释放很重要
- IWordRange是操作文档内容的核心接口
- 范围由起始和结束位置定义，可以精确控制
- 文本操作功能丰富，支持插入、删除、替换等操作

掌握文档结构和范围操作是进行复杂文档处理的基础，这些功能为后续的高级操作提供了必要的支持。