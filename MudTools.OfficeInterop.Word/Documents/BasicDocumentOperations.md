# 第2章：基本文档操作

在成功创建Word应用程序实例后，下一步就是学习如何操作文档。IWordApplication接口提供了丰富的属性和方法来管理Word应用程序及其文档。本章将详细介绍这些基本操作。

## IWordApplication接口详解

IWordApplication接口是MudTools.OfficeInterop.Word库中最重要的接口之一，它封装了Microsoft Word应用程序的各种功能。通过这个接口，我们可以控制Word应用程序的行为、访问文档集合、管理窗口等。

## 应用程序基础属性

IWordApplication接口提供了许多属性来控制和获取Word应用程序的状态：

```csharp
using var app = WordFactory.BlankWorkbook();
```

首先创建一个Word应用程序实例。

```csharp
// 设置应用程序窗口标题
app.Caption = "我的文档编辑器";
```

通过设置Caption属性来修改Word应用程序窗口的标题栏文本。

```csharp
// 控制状态栏显示
app.DisplayStatusBar = true;
```

DisplayStatusBar属性控制是否显示Word窗口底部的状态栏。

```csharp
// 控制滚动条显示
app.DisplayScrollBars = true;
```

DisplayScrollBars属性控制文档窗口中是否显示滚动条。

```csharp
// 获取应用程序可用区域尺寸
int width = app.UsableWidth;
int height = app.UsableHeight;
```

UsableWidth和UsableHeight属性返回Word文档窗口中可用于文档显示的最大宽度和高度（以磅为单位）。

```csharp
// 控制应用程序可见性
app.Visibility = WordAppVisibility.Visible;
```

通过Visibility属性控制Word应用程序窗口的可见性状态。

## 文档集合管理

通过[Documents](../Core/IWordApplication.cs#L74-L78)属性，我们可以访问所有打开的文档：

```csharp
using var app = WordFactory.BlankWorkbook();

// 获取文档集合
var documents = app.Documents;
```

Documents属性返回一个实现了[IWordDocuments](../Core/IWordDocuments.cs)接口的对象，代表所有已打开的文档集合。

```csharp
// 创建新文档
var newDoc = documents.Add();
```

使用Add方法创建一个新的空白文档并将其添加到文档集合中。

```csharp
// 打开现有文档
var existingDoc = documents.Open(@"C:\documents\example.docx");
```

使用Open方法打开一个已存在的文档文件。

```csharp
// 获取文档数量
int count = documents.Count;
```

Count属性返回当前打开的文档总数。

```csharp
// 遍历所有文档
for (int i = 1; i <= count; i++)
{
    var doc = documents.Item(i);
    Console.WriteLine($"文档 {i}: {doc.Name}");
}
```

通过索引（从1开始）访问文档集合中的特定文档，并输出文档名称。

## 活动文档和窗口管理

在任何时候，Word应用程序中都有一个活动文档和活动窗口：

```csharp
using var app = WordFactory.BlankWorkbook();

// 获取活动文档
var activeDoc = app.ActiveDocument;
```

ActiveDocument属性返回当前处于活动状态的文档对象。

```csharp
// 获取活动窗口
var activeWindow = app.ActiveWindow;
```

ActiveWindow属性返回当前活动的文档窗口对象。

```csharp
if (activeDoc != null)
{
    Console.WriteLine($"活动文档: {activeDoc.Name}");
}
```

检查活动文档是否存在并输出其名称。

```csharp
if (activeWindow != null)
{
    Console.WriteLine($"窗口标题: {activeWindow.Caption}");
}
```

检查活动窗口是否存在并输出其标题。

## 应用程序设置和选项

可以通过IWordApplication接口控制应用程序的各种设置：

```csharp
using var app = WordFactory.BlankWorkbook();

// 设置显示警告级别
app.DisplayAlerts = WdAlertLevel.None;
```

DisplayAlerts属性控制运行宏时显示的警告和消息级别。设置为None可以禁用所有警告对话框。

```csharp
// 控制自动完成提示
app.DisplayAutoCompleteTips = false;
```

DisplayAutoCompleteTips属性控制在键入时是否显示自动完成提示。

```csharp
// 控制屏幕提示显示
app.DisplayScreenTips = true;
```

DisplayScreenTips属性控制是否将批注、脚注、尾注和超链接显示为提示。

```csharp
// 设置取消键处理方式
app.EnableCancelKey = WdEnableCancelKey.Yes;
```

EnableCancelKey属性控制Word处理Ctrl+Break用户中断的方式。

```csharp
// 控制语言检查
app.CheckLanguage = true;
```

CheckLanguage属性控制Microsoft Word在键入时是否自动检测所使用的语言。

```csharp
// 控制屏幕更新
app.ScreenUpdating = true;
```

ScreenUpdating属性控制是否打开屏幕更新。在批量操作时设置为false可以提高性能。

## 显示和可见性控制

可以精确控制Word应用程序窗口的状态和显示方式：

```csharp
using var app = WordFactory.BlankWorkbook();

// 设置窗口状态（正常、最小化、最大化）
app.WordWindowState = WdWindowState.Maximize;
```

WordWindowState属性控制文档窗口或任务窗口的状态。

```csharp
// 控制应用程序可见性
app.Visibility = WordAppVisibility.Visible;
```

Visibility属性控制应用程序的可见性。

```csharp
// 获取或设置活动打印机
app.ActivePrinter = "Microsoft Print to PDF";
```

ActivePrinter属性获取或设置活动打印机的名称。

## 实际应用示例

以下是一个综合示例，展示如何使用这些基本操作：

```csharp
using MudTools.OfficeInterop;
using System;

class Program
{
    static void Main()
    {
        // 创建Word应用程序实例
        using var app = WordFactory.BlankWorkbook();
        
        try
        {
            // 配置应用程序
            app.Caption = "文档处理工具";
            app.DisplayStatusBar = true;
            app.Visibility = WordAppVisibility.Visible;
            app.WordWindowState = WdWindowState.Normal;
            
            // 获取活动文档
            var document = app.ActiveDocument;
            
            // 添加内容
            var range = document.Range();
            range.Text = "这是通过MudTools.OfficeInterop.Word创建的文档\n";
            
            // 保存文档
            document.SaveAs2(@"C:\temp\BasicDocumentExample.docx");
            
            Console.WriteLine($"文档已创建: {document.FullName}");
            Console.WriteLine($"文档名称: {document.Name}");
            Console.WriteLine($"文档路径: {document.Path}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"操作失败: {ex.Message}");
        }
        
        // 应用程序将在using语句结束时自动关闭
    }
}
```

让我们逐步分析这个示例：

```csharp
// 创建Word应用程序实例
using var app = WordFactory.BlankWorkbook();
```

使用BlankWorkbook方法创建一个新的Word应用程序实例。

```csharp
// 配置应用程序
app.Caption = "文档处理工具";
app.DisplayStatusBar = true;
app.Visibility = WordAppVisibility.Visible;
app.WordWindowState = WdWindowState.Normal;
```

配置应用程序的外观和行为。

```csharp
// 获取活动文档
var document = app.ActiveDocument;
```

获取当前活动的文档对象。

```csharp
// 添加内容
var range = document.Range();
range.Text = "这是通过MudTools.OfficeInterop.Word创建的文档\n";
```

获取文档范围并设置文本内容。

```csharp
// 保存文档
document.SaveAs2(@"C:\temp\BasicDocumentExample.docx");
```

将文档保存到指定路径。

## 应用场景

1. **文档管理系统**：通过管理多个文档，构建文档管理系统
2. **批量处理工具**：利用文档集合功能批量处理多个文档
3. **自动化办公**：通过设置应用程序选项，实现办公自动化
4. **集成开发环境**：在自定义IDE中嵌入Word文档编辑功能

## 要点总结

- IWordApplication接口是操作Word应用程序的核心接口
- 通过属性可以控制应用程序的外观和行为
- Documents属性提供了对文档集合的访问和管理能力
- 活动文档和窗口管理是多文档操作的基础
- 正确设置应用程序选项可以提升用户体验

掌握这些基本文档操作是进一步学习高级功能的基础，它们为复杂的文档处理任务提供了必要的支撑。