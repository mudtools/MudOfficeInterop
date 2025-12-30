# 第4章：选择区域操作

在Word文档处理中，选择区域(Selection)是一个重要概念，它代表当前光标位置或选中的文本区域。IWordSelection接口提供了丰富的功能来操作选择区域，模拟用户的实际操作。本章将详细介绍选择区域的操作方法。

## IWordSelection接口详解

IWordSelection接口封装了Word中选择区域的所有功能，允许开发者像用户一样操作文档。

```csharp
using var app = WordFactory.BlankDocument();
var selection = app.Selection;
```

通过Word应用程序实例获取当前的选择区域对象。

```csharp
if (selection != null)
{
    // 获取选择区域文本
    string selectedText = selection.Text;
    
    // 获取选择区域类型
    WdSelectionType selectionType = selection.Type;
    
    // 检查选择区域是否处于活动状态
    bool isActive = selection.Active;
}
```

检查选择区域对象是否存在，然后获取其文本内容、类型和活动状态。

## 选择区域的基本概念

选择区域是Word中一个动态的概念，它可以是一个插入点(光标位置)或选中的文本区域。

```csharp
using var app = WordFactory.BlankDocument();
var selection = app.Selection;

// 检查选择区域状态
if (selection != null)
{
    Console.WriteLine($"选择区域是否活动: {selection.Active}");
    Console.WriteLine($"选择区域类型: {selection.Type}");
    Console.WriteLine($"故事类型: {selection.StoryType}");
    Console.WriteLine($"故事长度: {selection.StoryLength}");
    
    // 检查是否处于行尾
    Console.WriteLine($"是否在行尾: {selection.IPAtEndOfLine}");
    Console.WriteLine($"是否在行末标记处: {selection.IsEndOfRowMark}");
}
```

输出选择区域的各种属性信息：
- Active：选择区域是否处于活动状态
- Type：选择区域的类型（插入点、文本选择等）
- StoryType：选择区域所在的故事类型（主文本、页眉、页脚等）
- StoryLength：当前故事的长度
- IPAtEndOfLine：插入点是否位于行尾
- IsEndOfRowMark：选择区域是否位于行末标记处

## 选择区域类型和属性

选择区域有不同的类型，每种类型有不同的操作特点：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;
var selection = app.Selection;

// 添加一些文本
document.Range().Text = "第一段文本。\n第二段文本。\n第三段文本。";

if (selection != null)
{
    // 检查选择区域类型
    switch (selection.Type)
    {
        case WdSelectionType.wdSelectionIP:
            Console.WriteLine("插入点选择");
            break;
        case WdSelectionType.wdSelectionNormal:
            Console.WriteLine("正常文本选择");
            break;
        case WdSelectionType.wdSelectionColumn:
            Console.WriteLine("列选择");
            break;
        case WdSelectionType.wdSelectionRow:
            Console.WriteLine("行选择");
            break;
        case WdSelectionType.wdSelectionBlock:
            Console.WriteLine("块选择");
            break;
    }
    
    // 设置选择模式
    selection.ExtendMode = false;
    selection.ColumnSelectMode = false;
}
```

不同选择区域类型的特点：
- wdSelectionIP：插入点选择，表示光标位置
- wdSelectionNormal：正常文本选择，表示选中的文本
- wdSelectionColumn：列选择，表示选中的列
- wdSelectionRow：行选择，表示选中的行
- wdSelectionBlock：块选择，表示选中的矩形块区域

## 文本选择和操作

选择区域提供了丰富的文本操作功能：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;
var selection = app.Selection;

// 添加内容
document.Range().Text = "这是第一段文本。\n这是第二段文本。\n这是第三段文本。";

if (selection != null)
{
    // 选择所有文本
    selection.WholeStory();
    
    // 获取选中的文本
    string allText = selection.Text;
    Console.WriteLine($"全部文本: {allText}");
    
    // 取消选择
    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
    
    // 选择特定范围
    selection.SetRange(0, 5);
    string selectedText = selection.Text;
    Console.WriteLine($"选中文本: {selectedText}");
    
    // 移动选择
    selection.MoveRight(WdUnits.wdCharacter, 1);
    selection.MoveDown(WdUnits.wdLine, 1);
}
```

逐步分析文本选择操作：
1. `WholeStory()`：选择整个文档内容
2. `Collapse(WdCollapseDirection.wdCollapseEnd)`：取消选择，将插入点移到末尾
3. `SetRange(0, 5)`：选择从位置0到位置5的文本
4. `MoveRight()`和`MoveDown()`：移动选择区域

## 选择区域的扩展和收缩

可以动态扩展或收缩选择区域：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;
var selection = app.Selection;

// 添加内容
document.Range().Text = "选择区域操作示例文本内容。";

if (selection != null)
{
    // 设置初始位置
    selection.Collapse(WdCollapseDirection.wdCollapseStart);
    
    // 扩展选择区域
    selection.MoveRight(WdUnits.wdWord, 2, WdMovementType.wdExtend);
    Console.WriteLine($"扩展后选中文本: '{selection.Text}'");
    
    // 进一步扩展
    selection.MoveRight(WdUnits.wdCharacter, 3, WdMovementType.wdExtend);
    Console.WriteLine($"再次扩展后选中文本: '{selection.Text}'");
    
    // 收缩选择区域
    selection.MoveLeft(WdUnits.wdWord, 1, WdMovementType.wdExtend);
    Console.WriteLine($"收缩后选中文本: '{selection.Text}'");
}
```

扩展和收缩操作详解：
1. `Collapse(WdCollapseDirection.wdCollapseStart)`：将选择区域折叠到开始位置
2. `MoveRight(WdUnits.wdWord, 2, WdMovementType.wdExtend)`：向右扩展选择2个单词
3. `MoveRight(WdUnits.wdCharacter, 3, WdMovementType.wdExtend)`：继续向右扩展选择3个字符
4. `MoveLeft(WdUnits.wdWord, 1, WdMovementType.wdExtend)`：向左收缩选择1个单词

## 高级选择操作

选择区域还支持更复杂的操作：

```csharp
using var app = WordFactory.BlankDocument();
var document = app.ActiveDocument;
var selection = app.Selection;

// 添加多段落内容
document.Range().Text = "第一段内容。\n第二段内容，包含更多文本。\n第三段内容。";

if (selection != null)
{
    // 选择整段
    selection.EndKey(WdUnits.wdStory);
    selection.HomeKey(WdUnits.wdLine, WdMovementType.wdExtend);
    Console.WriteLine($"整段文本: '{selection.Text}'");
    
    // 移动到文档开始
    selection.HomeKey(WdUnits.wdStory);
    
    // 选择到文档末尾
    selection.EndKey(WdUnits.wdStory, WdMovementType.wdExtend);
    Console.WriteLine($"全文本长度: {selection.Text.Length}");
    
    // 取消选择
    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
}
```

高级操作说明：
1. `EndKey(WdUnits.wdStory)`：移动到文档末尾
2. `HomeKey(WdUnits.wdLine, WdMovementType.wdExtend)`：从当前位置扩展选择到行首
3. `HomeKey(WdUnits.wdStory)`：移动到文档开始
4. `EndKey(WdUnits.wdStory, WdMovementType.wdExtend)`：从当前位置扩展选择到文档末尾

## 实际应用示例

以下示例演示了如何在实际应用中使用选择区域操作：

```csharp
using MudTools.OfficeInterop;
using System;

class SelectionDemo
{
    public static void DemonstrateSelectionOperations()
    {
        using var app = WordFactory.BlankDocument();
        app.Visible = true;
        
        try
        {
            var document = app.ActiveDocument;
            var selection = app.Selection;
            
            // 创建示例文档
            document.Range().Text = "Microsoft Word文档处理示例。\n" +
                                   "本示例演示选择区域操作。\n" +
                                   "包括文本选择、扩展和格式化。\n" +
                                   "这些操作模拟用户实际使用过程。";
            
            if (selection != null)
            {
                // 选择第一段
                selection.HomeKey(WdUnits.wdStory);
                selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
                
                // 格式化选中文本
                selection.Font.Bold = 1;
                selection.Font.Size = 14;
                selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                
                // 移动到下一段并选择
                selection.MoveDown(WdUnits.wdLine, 1);
                selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
                
                // 应用不同格式
                selection.Font.Bold = 0;
                selection.Font.Italic = 1;
                selection.Font.Color = WdColor.wdColorBlue;
                
                // 插入新内容
                selection.MoveDown(WdUnits.wdLine, 1);
                selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                selection.TypeText("\n这是通过选择区域插入的新内容。");
                
                // 保存文档
                document.SaveAs2(@"C:\temp\SelectionOperationsDemo.docx");
                
                Console.WriteLine("选择区域操作演示完成");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"操作失败: {ex.Message}");
        }
    }
}
```

逐步分析示例代码：

```csharp
// 创建示例文档
document.Range().Text = "Microsoft Word文档处理示例。\n" +
                       "本示例演示选择区域操作。\n" +
                       "包括文本选择、扩展和格式化。\n" +
                       "这些操作模拟用户实际使用过程。";
```

创建包含多段文本的示例文档。

```csharp
// 选择第一段
selection.HomeKey(WdUnits.wdStory);
selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
```

选择第一段文本：
1. `HomeKey(WdUnits.wdStory)`：移动到文档开始
2. `EndKey(WdUnits.wdLine, WdMovementType.wdExtend)`：扩展选择到行尾

```csharp
// 格式化选中文本
selection.Font.Bold = 1;
selection.Font.Size = 14;
selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

对选中文本应用格式：
1. 设置为粗体
2. 设置字体大小为14磅
3. 设置段落居中对齐

```csharp
// 移动到下一段并选择
selection.MoveDown(WdUnits.wdLine, 1);
selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
```

移动到下一段并选择该段文本。

```csharp
// 应用不同格式
selection.Font.Bold = 0;
selection.Font.Italic = 1;
selection.Font.Color = WdColor.wdColorBlue;
```

对第二段应用不同的格式：
1. 取消粗体
2. 设置为斜体
3. 设置字体颜色为蓝色

```csharp
// 插入新内容
selection.MoveDown(WdUnits.wdLine, 1);
selection.Collapse(WdCollapseDirection.wdCollapseEnd);
selection.TypeText("\n这是通过选择区域插入的新内容。");
```

在第三段后插入新内容：
1. 移动到第三段
2. 折叠选择区域到末尾
3. 插入新文本

## 应用场景

1. **文档编辑器**：模拟用户操作实现文档编辑功能
2. **格式化工具**：通过选择区域应用格式化样式
3. **内容查找替换**：结合查找功能定位和修改内容
4. **自动化脚本**：编写自动化处理文档的脚本

## 要点总结

- IWordSelection接口提供了对Word选择区域的完整访问能力
- 选择区域可以是插入点或选中的文本区域
- 不同类型的选择区域有不同的操作特点
- 可以通过移动、扩展和收缩来控制选择区域
- 选择区域操作模拟真实用户行为，功能强大
- 正确使用选择区域可以实现复杂的文档操作

掌握选择区域操作对于开发交互式文档处理应用非常重要，它是实现用户友好界面的关键技术之一。