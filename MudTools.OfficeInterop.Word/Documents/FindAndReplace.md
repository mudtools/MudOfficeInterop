# 第9章：查找和替换

查找和替换功能是Word文档处理中的重要工具，能够高效地修改文档内容。MudTools.OfficeInterop.Word库提供了强大的查找和替换功能，支持文本查找、格式查找、正则表达式等多种操作。本章将详细介绍如何使用这些功能进行批量文本处理。

## 查找功能详解

查找功能允许在文档中定位特定内容，可以基于文本、格式或其他属性进行查找。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
document.Range().Text = "这是示例文本。\n查找和替换功能演示。\n示例文本包含多个实例。";

// 获取查找对象
var find = document.Range().Find;
```

通过Range().Find属性获取查找对象，用于执行查找操作。

```csharp
// 基本文本查找
find.ClearFormatting();
find.Text = "示例";
find.Forward = true;
find.Wrap = WdFindWrap.wdFindContinue;
```

设置基本查找参数：
- ClearFormatting()：清除之前的查找格式设置
- Text：要查找的文本内容
- Forward = true：向前查找（从当前位置向文档末尾）
- Wrap = WdFindWrap.wdFindContinue：查找到文档末尾后继续从开头查找

```csharp
// 执行查找
bool found = find.Execute();

if (found)
{
    Console.WriteLine("找到了文本 '示例'");
}
else
{
    Console.WriteLine("未找到文本 '示例'");
}

// 查找下一个匹配项
while (find.Execute())
{
    Console.WriteLine("找到下一个匹配项");
}
```

执行查找操作：
- Execute()：执行查找，返回bool值表示是否找到
- 循环执行可以找到所有匹配项

## 替换操作

替换功能可以在查找到内容后进行替换操作，支持简单的文本替换和复杂的格式替换。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
document.Range().Text = "原文本1\n原文本2\n原文本3\n";

// 获取查找和替换对象
var find = document.Range().Find;
var replace = find; // 替换对象与查找对象是同一个
```

获取查找对象，替换操作使用相同的对象。

```csharp
// 设置查找和替换参数
find.ClearFormatting();
replace.ClearFormatting();
find.Text = "原文本";
replace.Text = "新文本";
```

设置查找和替换文本内容。

```csharp
// 执行替换（只替换第一个匹配项）
find.Execute(
    FindText: "原文本",
    ReplaceWith: "新文本",
    Replace: WdReplace.wdReplaceOne
);
```

执行单次替换：
- FindText：要查找的文本
- ReplaceWith：替换后的文本
- Replace: WdReplace.wdReplaceOne：只替换第一个匹配项

```csharp
// 执行全部替换
find.Execute(
    FindText: "原文本",
    ReplaceWith: "新文本",
    Replace: WdReplace.wdReplaceAll
);
```

执行全部替换：
- Replace: WdReplace.wdReplaceAll：替换所有匹配项

## 格式查找和替换

除了文本查找，还可以基于格式进行查找和替换。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
var range = document.Range();
range.Text = "普通文本\n粗体文本\n斜体文本\n";

// 设置粗体文本
var boldRange = document.Range(6, 10); // "粗体文本"
boldRange.Font.Bold = 1;

// 设置斜体文本
var italicRange = document.Range(11, 15); // "斜体文本"
italicRange.Font.Italic = 1;
```

创建包含不同格式的示例文本。

```csharp
// 查找粗体文本
var find = document.Range().Find;
find.ClearFormatting();
find.Font.Bold = 1; // 查找粗体文本
find.Text = ""; // 文本可以为空，只基于格式查找

// 执行查找
bool found = find.Execute();
if (found)
{
    Console.WriteLine("找到了粗体文本");
}
```

基于格式查找：
- Font.Bold = 1：查找粗体文本
- Text为空：只基于格式查找

```csharp
// 替换粗体格式为下划线格式
find.ClearFormatting();
find.Font.Bold = 1;
find.Replacement.ClearFormatting();
find.Replacement.Font.Underline = WdUnderline.wdUnderlineSingle;

find.Execute(
    FindText: "",
    ReplaceWith: "",
    Replace: WdReplace.wdReplaceAll
);
```

将粗体格式替换为下划线格式：
- 查找粗体文本
- 设置替换格式为下划线
- 执行全部替换

## 正则表达式支持

Word的查找功能支持类似正则表达式的通配符模式。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
document.Range().Text = "电话: 138-1234-5678\n邮箱: example@test.com\n日期: 2025-10-06\n";

// 使用通配符查找电话号码
var find = document.Range().Find;
find.ClearFormatting();
find.Text = "[0-9]{3}-[0-9]{4}-[0-9]{4}"; // 电话号码模式
find.MatchWildcards = true;
```

启用通配符匹配并设置电话号码模式：
- [0-9]{3}：匹配3个数字
- -：匹配连字符
- [0-9]{4}：匹配4个数字

```csharp
bool found = find.Execute();
if (found)
{
    Console.WriteLine("找到了电话号码");
}
```

执行查找电话号码。

```csharp
// 使用通配符查找邮箱
find.Text = "[a-zA-Z0-9]*@[a-zA-Z0-9]*\.[a-zA-Z]*";
find.MatchWildcards = true;

found = find.Execute();
if (found)
{
    Console.WriteLine("找到了邮箱地址");
}
```

查找邮箱地址模式：
- [a-zA-Z0-9]*：匹配字母和数字
- @：匹配@符号
- \.：匹配点号（需要转义）

```csharp
// 使用通配符查找日期
find.Text = "[0-9]{4}-[0-9]{2}-[0-9]{2}";
find.MatchWildcards = true;

found = find.Execute();
if (found)
{
    Console.WriteLine("找到了日期");
}
```

查找日期格式：YYYY-MM-DD。

## 高级查找选项

查找功能支持多种高级选项，如大小写敏感、全字匹配等。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
document.Range().Text = "Word word WORD\nText text TEXT\n";

var find = document.Range().Find;
find.ClearFormatting();

// 大小写敏感查找
find.Text = "Word";
find.MatchCase = true;
bool found1 = find.Execute();
Console.WriteLine($"大小写敏感查找: {found1}");
```

启用大小写敏感查找。

```csharp
// 全字匹配查找
find.Text = "word";
find.MatchCase = false;
find.MatchWholeWord = true;
bool found2 = find.Execute();
Console.WriteLine($"全字匹配查找: {found2}");
```

启用全字匹配查找。

```csharp
// 使用同义词库查找
find.Text = "car";
find.MatchFuzzy = true;
bool found3 = find.Execute();
Console.WriteLine($"同义词查找: {found3}");
```

启用同义词模糊查找。

```csharp
// 向前查找
find.Text = "word";
find.Forward = true;
find.Wrap = WdFindWrap.wdFindStop;
bool found4 = find.Execute();
Console.WriteLine($"向前查找: {found4}");
```

设置查找方向和换行行为。

## 批量文本处理

结合查找和替换功能，可以实现复杂的批量文本处理。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 添加示例内容
document.Range().Text = "Mr. Zhang\nMrs. Li\nDr. Wang\nMr. Liu\nMs. Chen\n";

// 批量替换称谓
var find = document.Range().Find;

// 替换 "Mr." 为 "先生"
find.Execute(
    FindText: "Mr.",
    ReplaceWith: "先生",
    Replace: WdReplace.wdReplaceAll
);
```

批量替换称谓的第一个示例。

```csharp
// 替换 "Mrs." 为 "夫人"
find.Execute(
    FindText: "Mrs.",
    ReplaceWith: "夫人",
    Replace: WdReplace.wdReplaceAll
);

// 替换 "Dr." 为 "博士"
find.Execute(
    FindText: "Dr.",
    ReplaceWith: "博士",
    Replace: WdReplace.wdReplaceAll
);

// 替换 "Ms." 为 "女士"
find.Execute(
    FindText: "Ms.",
    ReplaceWith: "女士",
    Replace: WdReplace.wdReplaceAll
);

Console.WriteLine("称谓替换完成");
```

完成所有称谓的替换。

## 实际应用示例

以下示例演示了如何创建一个文档清理工具，批量处理文档中的各种内容：

```csharp
using MudTools.OfficeInterop;
using System;
using System.Collections.Generic;

class DocumentCleanupTool
{
    public static void CleanupDocument()
    {
        using var app = WordFactory.Open(@"C:\temp\DocumentToClean.docx");
        var document = app.ActiveDocument;
        
        try
        {
            Console.WriteLine("开始文档清理...");
            
            // 1. 清理多余的空格
            CleanupExtraSpaces(document);
            
            // 2. 标准化称谓
            StandardizeTitles(document);
            
            // 3. 清理空白行
            RemoveExtraBlankLines(document);
            
            // 4. 标准化日期格式
            StandardizeDateFormats(document);
            
            // 5. 更新文档属性
            UpdateDocumentProperties(document);
            
            // 保存清理后的文档
            document.SaveAs2(@"C:\temp\CleanedDocument.docx");
            
            Console.WriteLine($"文档清理完成: {document.FullName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"清理文档时出错: {ex.Message}");
        }
    }
```

文档清理工具主函数。

```csharp
    private static void CleanupExtraSpaces(var document)
    {
        var find = document.Range().Find;
        
        // 替换多个空格为单个空格
        find.Execute(
            FindText: "  ", // 两个空格
            ReplaceWith: " ", // 一个空格
            Replace: WdReplace.wdReplaceAll
        );
```

清理多余空格。

```csharp
        // 清理行首空格
        find.Execute(
            FindText: "^p ", // 段落标记后跟空格
            ReplaceWith: "^p", // 仅段落标记
            Replace: WdReplace.wdReplaceAll
        );
        
        Console.WriteLine("多余空格清理完成");
    }
```

清理行首空格，^p表示段落标记。

```csharp
    private static void StandardizeTitles(var document)
    {
        var find = document.Range().Find;
        
        // 标准化公司名称
        var companyReplacements = new Dictionary<string, string>
        {
            {"某某公司", "ABC有限公司"},
            {"XYZ集团", "XYZ集团股份有限公司"},
            {"DEF企业", "DEF企业发展有限公司"}
        };
        
        foreach (var pair in companyReplacements)
        {
            find.Execute(
                FindText: pair.Key,
                ReplaceWith: pair.Value,
                Replace: WdReplace.wdReplaceAll
            );
        }
        
        Console.WriteLine("称谓标准化完成");
    }
```

标准化公司名称。

```csharp
    private static void RemoveExtraBlankLines(var document)
    {
        var find = document.Range().Find;
        
        // 删除连续的空行（保留一个）
        find.Execute(
            FindText: "^p^p^p", // 三个连续段落标记
            ReplaceWith: "^p^p", // 两个段落标记
            Replace: WdReplace.wdReplaceAll
        );
```

删除多余空行。

```csharp
        // 再次执行以处理更多连续空行
        find.Execute(
            FindText: "^p^p^p",
            ReplaceWith: "^p^p",
            Replace: WdReplace.wdReplaceAll
        );
        
        Console.WriteLine("空白行清理完成");
    }
```

再次执行确保清理彻底。

```csharp
    private static void StandardizeDateFormats(var document)
    {
        var find = document.Range().Find;
        
        // 使用通配符查找并标准化日期格式
        find.MatchWildcards = true;
        
        // 查找 YYYY/MM/DD 格式并替换为 YYYY-MM-DD
        find.Execute(
            FindText: "([0-9]{4})/([0-9]{2})/([0-9]{2})",
            ReplaceWith: "\\1-\\2-\\3",
            Replace: WdReplace.wdReplaceAll
        );
```

标准化日期格式，使用捕获组。

```csharp
        // 查找 YYYY.MM.DD 格式并替换为 YYYY-MM-DD
        find.Execute(
            FindText: "([0-9]{4})\.([0-9]{2})\.([0-9]{2})",
            ReplaceWith: "\\1-\\2-\\3",
            Replace: WdReplace.wdReplaceAll
        );
        
        find.MatchWildcards = false;
        Console.WriteLine("日期格式标准化完成");
    }
```

处理点号分隔的日期格式。

```csharp
    private static void UpdateDocumentProperties(var document)
    {
        // 更新文档属性
        document.Title = "清理后的文档";
        document.Author = "文档清理工具";
        document.Subject = "已清理的文档";
        document.Keywords = "清理, 标准化, 自动化";
        
        Console.WriteLine("文档属性更新完成");
    }
}
```

更新文档属性。

## 应用场景

1. **文档标准化**：批量修改公司文档中的标准术语、格式等
2. **模板处理**：处理文档模板中的占位符
3. **数据清理**：清理从其他系统导入的文档数据
4. **版本更新**：批量更新文档中的版本信息、日期等

## 要点总结

- 查找功能支持基于文本和格式的查找操作
- 替换功能可以在查找到内容后进行替换
- 支持通配符模式，实现类似正则表达式的查找
- 提供多种高级查找选项，如大小写敏感、全字匹配等
- 可以实现复杂的批量文本处理任务
- 结合使用可以创建强大的文档处理工具

掌握查找和替换功能对于高效处理Word文档至关重要，这些功能使开发者能够自动化完成大量重复性的文本编辑工作。