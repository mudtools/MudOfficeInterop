# 第10章：邮件合并

邮件合并是Word中一个强大的功能，可以将数据源中的信息与文档模板结合，批量生成个性化的文档。MudTools.OfficeInterop.Word库提供了完整的邮件合并功能支持，包括数据源连接、字段操作、执行合并等。本章将详细介绍如何使用这些功能批量生成个性化文档。

## 邮件合并基础

邮件合并涉及三个主要组件：主文档（包含合并字段的模板）、数据源（包含实际数据）和合并结果（生成的个性化文档）。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;

// 设置文档为邮件合并主文档
var mailMerge = document.MailMerge;
```

通过Document.MailMerge属性获取邮件合并对象。

```csharp
// 检查文档是否为邮件合并主文档
bool isMailMergeDoc = mailMerge.MainDocumentType != WdMailMergeMainDocType.wdNotAMergeDocument;
Console.WriteLine($"是否为邮件合并文档: {isMailMergeDoc}");
```

检查文档是否已设置为邮件合并主文档。

```csharp
// 设置主文档类型
mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters; // 信函
// mailMerge.MainDocumentType = WdMailMergeMainDocType.wdMailingLabels; // 邮寄标签
// mailMerge.MainDocumentType = WdMailMergeMainDocType.wdCatalog; // 目录
// mailMerge.MainDocumentType = WdMailMergeMainDocType.wdEnvelopes; // 信封
// mailMerge.MainDocumentType = WdMailMergeMainDocType.wdDirectory; // 名录
```

设置邮件合并主文档类型：
- wdFormLetters：信函（最常用）
- wdMailingLabels：邮寄标签
- wdCatalog：目录
- wdEnvelopes：信封
- wdDirectory：名录

## 数据源连接

邮件合并需要连接到数据源，可以是数据库、Excel文件、文本文件等。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;
var mailMerge = document.MailMerge;

// 连接Excel数据源
try
{
    mailMerge.OpenDataSource(
        Name: @"C:\data\Customers.xlsx",
        Connection: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\data\\Customers.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"",
        SQLStatement: "SELECT * FROM [Sheet1$]"
    );
```

连接Excel数据源：
- Name：数据源文件路径
- Connection：OLEDB连接字符串
- SQLStatement：SQL查询语句

```csharp
    Console.WriteLine("数据源连接成功");
}
catch (Exception ex)
{
    Console.WriteLine($"数据源连接失败: {ex.Message}");
}
```

处理连接异常。

```csharp
// 连接文本文件数据源
try
{
    mailMerge.OpenDataSource(
        Name: @"C:\data\Customers.csv",
        Connection: "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\data\\;Extended Properties=\"text;HDR=YES;FMT=Delimited\"",
        SQLStatement: "SELECT * FROM Customers.csv"
    );
```

连接CSV文本文件数据源。

```csharp
    Console.WriteLine("CSV数据源连接成功");
}
catch (Exception ex)
{
    Console.WriteLine($"CSV数据源连接失败: {ex.Message}");
}
```

## 合并字段操作

合并字段是邮件合并的核心，它们在主文档中作为占位符，会被数据源中的实际数据替换。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;
var mailMerge = document.MailMerge;

// 添加合并字段到文档
var range = document.Range();

// 添加标题
range.Text = "客户信息\n\n";
range.Font.Bold = 1;
range.Font.Size = 16;
range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

添加文档标题并设置格式。

```csharp
// 移动到文档末尾
range.Collapse(WdCollapseDirection.wdCollapseEnd);

// 添加包含合并字段的内容
range.Text = "尊敬的 ";
range.Collapse(WdCollapseDirection.wdCollapseEnd);

// 插入合并字段
mailMerge.Fields.Add(range, "姓名");
range.Text = " 先生/女士：\n\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);
```

插入合并字段：
- Collapse(WdCollapseDirection.wdCollapseEnd)：将光标移到末尾
- Fields.Add(range, "姓名")：插入名为"姓名"的合并字段

```csharp
range.Text = "您的客户编号是：";
mailMerge.Fields.Add(range, "客户编号");
range.Text = "\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);

range.Text = "联系电话：";
mailMerge.Fields.Add(range, "电话");
range.Text = "\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);

range.Text = "电子邮箱：";
mailMerge.Fields.Add(range, "邮箱");
range.Text = "\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);

range.Text = "地址：";
mailMerge.Fields.Add(range, "地址");
range.Text = "\n\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);
```

继续添加其他合并字段。

```csharp
range.Text = "感谢您选择我们的服务！\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);

// 查看所有合并字段
Console.WriteLine($"合并字段数量: {mailMerge.Fields.Count}");
for (int i = 1; i <= mailMerge.Fields.Count; i++)
{
    Console.WriteLine($"字段 {i}: {mailMerge.Fields.Item(i).Code}");
}
```

显示所有合并字段信息。

## 执行邮件合并

连接数据源并设置好合并字段后，就可以执行邮件合并操作。

```csharp
using var app = WordFactory.CreateFrom(@"C:\templates\LetterTemplate.dotx");
var document = app.ActiveDocument;
var mailMerge = document.MailMerge;

try
{
    // 连接数据源
    mailMerge.OpenDataSource(
        Name: @"C:\data\Customers.xlsx",
        Connection: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\data\\Customers.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"",
        SQLStatement: "SELECT * FROM [Sheet1$]"
    );
```

打开数据源。

```csharp
    // 设置主文档类型
    mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
    
    // 执行邮件合并到新文档
    mailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
    mailMerge.Execute(Pause: false);
```

执行邮件合并：
- Destination = WdMailMergeDestination.wdSendToNewDocument：将结果发送到新文档
- Execute(Pause: false)：执行合并，不暂停

```csharp
    Console.WriteLine("邮件合并执行完成");
}
catch (Exception ex)
{
    Console.WriteLine($"邮件合并执行失败: {ex.Message}");
}
```

## 高级邮件合并操作

邮件合并支持一些高级功能，如条件合并、计算字段等。

```csharp
using var app = WordFactory.BlankWorkbook();
var document = app.ActiveDocument;
var mailMerge = document.MailMerge;

// 添加条件合并字段
var range = document.Range();
range.Text = "亲爱的客户：\n\n";

// 添加条件文本（根据性别）
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "{ IF { MERGEFIELD 性别 } = \"男\" \"先生\" \"女士\" }";
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "\n\n";
```

添加条件字段，根据性别显示不同称谓。

```csharp
// 添加计算字段（计算折扣）
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "您的订单总额为：{ MERGEFIELD 订单金额 } 元\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);
range.Text = "享受折扣后价格为：{ = { MERGEFIELD 订单金额 } * 0.9 } 元\n\n";
range.Collapse(WdCollapseDirection.wdCollapseEnd);
```

添加计算字段，计算折扣价格。

```csharp
// 添加日期字段
range.Text = "生成日期：{ DATE \\@ \"yyyy年MM月dd日\" }\n\n";

// 更新所有字段
document.Range().Fields.Update();
```

添加日期字段并更新所有字段。

## 实际应用示例

以下示例演示了如何创建一个完整的邮件合并系统：

```csharp
using MudTools.OfficeInterop;
using System;
using System.Data;
using System.Data.OleDb;

class MailMergeSystem
{
    public static void GenerateCustomerLetters()
    {
        try
        {
            // 1. 创建或准备数据源（这里使用内存数据模拟）
            CreateSampleDataSource();
            
            // 2. 创建邮件合并模板
            CreateMailMergeTemplate();
            
            // 3. 执行邮件合并
            ExecuteMailMerge();
            
            Console.WriteLine("客户信函生成完成");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"邮件合并过程中出错: {ex.Message}");
        }
    }
```

邮件合并系统主函数。

```csharp
    private static void CreateSampleDataSource()
    {
        // 在实际应用中，这里会连接到真实的数据库或文件
        // 为了演示，我们创建一个简单的Excel文件
        
        Console.WriteLine("创建示例数据源...");
        // 这里可以使用Excel.Interop或其他库创建Excel文件
        // 为简化示例，我们假设数据源已存在
    }
    
    private static void CreateMailMergeTemplate()
    {
        using var app = WordFactory.BlankWorkbook();
        var document = app.ActiveDocument;
        var mailMerge = document.MailMerge;
        
        Console.WriteLine("创建邮件合并模板...");
        
        // 设置文档为邮件合并主文档
        mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
```

创建邮件合并模板。

```csharp
        // 添加模板内容
        var range = document.Range();
        
        // 页眉
        range.Text = "ABC有限公司\n地址：某某市某某区某某路123号\n电话：010-12345678\n\n";
        range.Font.Bold = 1;
        range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        
        // 日期
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "{ DATE \\@ \"yyyy年MM月dd日\" }\n\n";
        range.Font.Bold = 0;
        range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
```

添加模板页眉和日期。

```csharp
        // 收件人地址
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "{ MERGEFIELD 客户姓名 }\n";
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "{ MERGEFIELD 地址 }\n\n";
        
        // 称呼
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "尊敬的 { MERGEFIELD 客户姓名 } 先生/女士：\n\n";
        
        // 正文
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "感谢您一直以来对我们公司的支持与信任。我们很高兴地通知您，您的账户信息已更新。\n\n";
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "以下是您的账户信息：\n\n";
```

添加模板正文内容。

```csharp
        // 账户信息表格
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "客户编号：\t{ MERGEFIELD 客户编号 }\n";
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "账户余额：\t{ MERGEFIELD 账户余额 } 元\n";
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "信用等级：\t{ MERGEFIELD 信用等级 }\n\n";
        
        // 条款和条件
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "如有任何疑问，请随时与我们联系。\n\n";
        
        // 结尾
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
        range.Text = "此致\n敬礼！\n\nABC有限公司客户服务部\n";
```

添加模板结尾内容。

```csharp
        // 更新字段
        document.Range().Fields.Update();
        
        // 保存模板
        document.SaveAs2(@"C:\temp\CustomerLetterTemplate.dotx");
        Console.WriteLine("邮件合并模板已创建: CustomerLetterTemplate.dotx");
    }
```

更新字段并保存模板。

```csharp
    private static void ExecuteMailMerge()
    {
        using var app = WordFactory.CreateFrom(@"C:\temp\CustomerLetterTemplate.dotx");
        var document = app.ActiveDocument;
        var mailMerge = document.MailMerge;
        
        Console.WriteLine("执行邮件合并...");
        
        try
        {
            // 连接数据源（这里假设有一个Excel文件）
            mailMerge.OpenDataSource(
                Name: @"C:\temp\CustomerData.xlsx",
                Connection: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\temp\\CustomerData.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"",
                SQLStatement: "SELECT * FROM [Customers$]"
            );
```

执行邮件合并。

```csharp
            // 设置主文档类型
            mailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;
            
            // 执行邮件合并到新文档
            mailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
            mailMerge.Execute(Pause: false);
            
            // 保存合并结果
            var resultDoc = app.ActiveDocument;
            resultDoc.SaveAs2(@"C:\temp\CustomerLettersResult.docx");
            
            Console.WriteLine("邮件合并结果已保存: CustomerLettersResult.docx");
        }
        catch (System.IO.FileNotFoundException)
        {
            Console.WriteLine("警告：未找到数据源文件，使用示例数据执行合并");
```

处理数据源文件不存在的情况。

```csharp
            // 使用示例数据执行合并
            mailMerge.DataSource.DataFields.Add("客户姓名", "张三");
            mailMerge.DataSource.DataFields.Add("地址", "北京市朝阳区某某街道123号");
            mailMerge.DataSource.DataFields.Add("客户编号", "C12345");
            mailMerge.DataSource.DataFields.Add("账户余额", "10000");
            mailMerge.DataSource.DataFields.Add("信用等级", "A");
            
            // 执行合并
            mailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
            mailMerge.Execute(Pause: false);
            
            // 保存结果
            var resultDoc = app.ActiveDocument;
            resultDoc.SaveAs2(@"C:\temp\CustomerLettersSample.docx");
            Console.WriteLine("示例邮件合并结果已保存: CustomerLettersSample.docx");
        }
    }
}
```

使用示例数据执行合并并保存结果。

## 应用场景

1. **批量生成邀请函**：为活动参与者批量生成个性化邀请函
2. **制作报告**：为不同客户生成个性化的财务报告或分析报告
3. **生成证书**：为培训参与者批量生成结业证书
4. **客户通信**：向客户发送个性化的信函、通知等

## 要点总结

- 邮件合并涉及主文档、数据源和合并结果三个核心组件
- 支持连接多种数据源，包括数据库、Excel文件、文本文件等
- 合并字段作为占位符，在合并时被实际数据替换
- 提供条件合并、计算字段等高级功能
- 可以将结果发送到新文档、打印机或电子邮件
- 适用于各种批量文档生成场景

掌握邮件合并功能对于自动化生成大量个性化文档至关重要，这些功能使开发者能够高效地处理批量文档生成任务。