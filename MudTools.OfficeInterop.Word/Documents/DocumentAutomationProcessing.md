# 第15章：文档自动化处理

文档自动化处理是提高办公效率的重要手段，通过MudTools.OfficeInterop.Word库，我们可以实现文档的批量处理、格式转换、自动化工作流等功能。本章将详细介绍如何构建完整的文档自动化处理系统。

## 批量文档处理

批量文档处理可以同时对多个文档执行相同的操作，大大提高工作效率。

```csharp
using MudTools.OfficeInterop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

class BatchDocumentProcessor
{
    public static void ProcessDocumentsInBatch(string inputDirectory, string outputDirectory, 
                                              string filePattern = "*.docx")
    {
        try
        {
            // 确保输出目录存在
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
```

确保输出目录存在，如果不存在则创建。

```csharp
            // 获取所有匹配的文件
            var files = Directory.GetFiles(inputDirectory, filePattern);
            
            Console.WriteLine($"找到 {files.Length} 个文档需要处理");
            
            int processedCount = 0;
            int errorCount = 0;
```

获取指定目录下所有匹配模式的文件。

```csharp
            foreach (var file in files)
            {
                try
                {
                    Console.WriteLine($"正在处理: {Path.GetFileName(file)}");
                    
                    // 处理单个文档
                    ProcessSingleDocument(file, outputDirectory);
                    
                    processedCount++;
                    Console.WriteLine($"处理完成: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"处理 {Path.GetFileName(file)} 时出错: {ex.Message}");
                }
            }
```

遍历所有文件并处理每个文档。

```csharp
            Console.WriteLine($"\n批量处理完成:");
            Console.WriteLine($"成功处理: {processedCount} 个文档");
            Console.WriteLine($"处理失败: {errorCount} 个文档");
            Console.WriteLine($"总计处理: {files.Length} 个文档");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"批量处理过程中出错: {ex.Message}");
        }
    }
```

输出处理结果统计。

```csharp
    private static void ProcessSingleDocument(string inputFilePath, string outputDirectory)
    {
        using var app = WordFactory.Open(inputFilePath);
        var document = app.ActiveDocument;
        
        // 执行文档处理操作
        // 例如：标准化格式、更新字段、添加页眉页脚等
        StandardizeDocumentFormat(document);
        UpdateDocumentFields(document);
        AddHeaderFooter(document);
```

处理单个文档，执行各种操作。

```csharp
        // 生成输出文件路径
        var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
        var outputFilePath = Path.Combine(outputDirectory, $"{fileName}_processed.docx");
        
        // 保存处理后的文档
        document.SaveAs2(outputFilePath);
    }
```

保存处理后的文档。

```csharp
    private static void StandardizeDocumentFormat(var document)
    {
        // 标准化字体
        var range = document.Range();
        range.Font.Name = "宋体";
        range.Font.Size = 12;
        
        // 标准化段落格式
        foreach (var paragraph in document.Paragraphs)
        {
            paragraph.Format.LineSpacing = 1.5f; // 1.5倍行距
            paragraph.Format.SpaceAfter = 12;    // 段后间距
        }
        
        Console.WriteLine("  - 文档格式已标准化");
    }
```

标准化文档格式。

```csharp
    private static void UpdateDocumentFields(var document)
    {
        // 更新所有字段
        document.Range().Fields.Update();
        Console.WriteLine("  - 文档字段已更新");
    }
    
    private static void AddHeaderFooter(var document)
    {
        // 添加页眉
        var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        headerRange.Text = "公司文档";
        headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

添加页眉内容。

```csharp
        // 添加页脚（包含页码）
        var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
        footerRange.Text = " 第 页";
        footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        
        Console.WriteLine("  - 页眉页脚已添加");
    }
}
```

添加页脚内容。

## 文档转换

文档转换功能可以将文档从一种格式转换为另一种格式。

```csharp
class DocumentConverter
{
    public static void ConvertDocuments(string inputDirectory, string outputDirectory, 
                                      WdSaveFormat targetFormat, string filePattern = "*.doc")
    {
        try
        {
            // 确保输出目录存在
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
```

确保输出目录存在。

```csharp
            // 获取所有匹配的文件
            var files = Directory.GetFiles(inputDirectory, filePattern);
            
            Console.WriteLine($"找到 {files.Length} 个文档需要转换");
            
            int convertedCount = 0;
            int errorCount = 0;
```

获取需要转换的文件列表。

```csharp
            foreach (var file in files)
            {
                try
                {
                    Console.WriteLine($"正在转换: {Path.GetFileName(file)}");
                    
                    // 转换单个文档
                    ConvertSingleDocument(file, outputDirectory, targetFormat);
                    
                    convertedCount++;
                    Console.WriteLine($"转换完成: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"转换 {Path.GetFileName(file)} 时出错: {ex.Message}");
                }
            }
```

遍历并转换每个文档。

```csharp
            Console.WriteLine($"\n批量转换完成:");
            Console.WriteLine($"成功转换: {convertedCount} 个文档");
            Console.WriteLine($"转换失败: {errorCount} 个文档");
            Console.WriteLine($"总计转换: {files.Length} 个文档");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"批量转换过程中出错: {ex.Message}");
        }
    }
```

输出转换结果统计。

```csharp
    private static void ConvertSingleDocument(string inputFilePath, string outputDirectory, 
                                            WdSaveFormat targetFormat)
    {
        using var app = WordFactory.Open(inputFilePath);
        var document = app.ActiveDocument;
        
        // 生成输出文件路径
        var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
        string extension = GetExtensionForFormat(targetFormat);
        var outputFilePath = Path.Combine(outputDirectory, $"{fileName}{extension}");
```

生成输出文件路径。

```csharp
        // 保存为指定格式
        document.SaveAs2(outputFilePath, targetFormat);
        
        Console.WriteLine($"  - 已转换为: {extension}");
    }
```

保存为指定格式。

```csharp
    private static string GetExtensionForFormat(WdSaveFormat format)
    {
        return format switch
        {
            WdSaveFormat.wdFormatDocument => ".doc",
            WdSaveFormat.wdFormatXMLDocument => ".xml",
            WdSaveFormat.wdFormatPDF => ".pdf",
            WdSaveFormat.wdFormatRTF => ".rtf",
            WdSaveFormat.wdFormatFilteredHTML => ".htm",
            WdSaveFormat.wdFormatHTML => ".html",
            _ => ".docx"
        };
    }
```

根据格式获取文件扩展名。

```csharp
    // 特殊转换示例：Word到PDF
    public static void ConvertToPdf(string inputDirectory, string outputDirectory)
    {
        ConvertDocuments(inputDirectory, outputDirectory, WdSaveFormat.wdFormatPDF, "*.docx");
    }
    
    // 特殊转换示例：Word到HTML
    public static void ConvertToHtml(string inputDirectory, string outputDirectory)
    {
        ConvertDocuments(inputDirectory, outputDirectory, WdSaveFormat.wdFormatFilteredHTML, "*.docx");
    }
}
```

提供特定格式转换方法。

## 自动化工作流

自动化工作流可以将多个文档处理步骤组合成一个完整的处理流程。

```csharp
class DocumentAutomationWorkflow
{
    public class WorkflowConfiguration
    {
        public bool StandardizeFormat { get; set; } = true;
        public bool UpdateFields { get; set; } = true;
        public bool AddHeaderFooter { get; set; } = true;
        public bool ConvertToPdf { get; set; } = false;
        public bool GenerateTableOfContents { get; set; } = false;
        public string WatermarkText { get; set; } = null;
    }
```

定义工作流配置类。

```csharp
    public static void ExecuteWorkflow(string inputDirectory, string outputDirectory, 
                                     WorkflowConfiguration config)
    {
        try
        {
            Console.WriteLine("=== 开始执行文档自动化工作流 ===");
            
            // 确保输出目录存在
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
```

执行自动化工作流。

```csharp
            // 获取所有Word文档
            var docFiles = Directory.GetFiles(inputDirectory, "*.doc");
            var docxFiles = Directory.GetFiles(inputDirectory, "*.docx");
            var allFiles = docFiles.Concat(docxFiles).ToArray();
            
            Console.WriteLine($"找到 {allFiles.Length} 个文档需要处理");
            
            int processedCount = 0;
            int errorCount = 0;
```

获取所有需要处理的文档。

```csharp
            foreach (var file in allFiles)
            {
                try
                {
                    Console.WriteLine($"\n正在处理: {Path.GetFileName(file)}");
                    
                    // 执行工作流
                    ExecuteSingleDocumentWorkflow(file, outputDirectory, config);
                    
                    processedCount++;
                    Console.WriteLine($"处理完成: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"处理 {Path.GetFileName(file)} 时出错: {ex.Message}");
                }
            }
```

处理每个文档。

```csharp
            Console.WriteLine($"\n=== 工作流执行完成 ===");
            Console.WriteLine($"成功处理: {processedCount} 个文档");
            Console.WriteLine($"处理失败: {errorCount} 个文档");
            Console.WriteLine($"总计处理: {allFiles.Length} 个文档");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"工作流执行过程中出错: {ex.Message}");
        }
    }
```

输出工作流执行结果。

```csharp
    private static void ExecuteSingleDocumentWorkflow(string inputFilePath, string outputDirectory, 
                                                    WorkflowConfiguration config)
    {
        using var app = WordFactory.Open(inputFilePath);
        var document = app.ActiveDocument;
        
        // 执行配置的处理步骤
        if (config.StandardizeFormat)
        {
            StandardizeDocumentFormat(document);
        }
```

根据配置执行处理步骤。

```csharp
        if (config.UpdateFields)
        {
            UpdateDocumentFields(document);
        }
        
        if (config.AddHeaderFooter)
        {
            AddHeaderFooter(document);
        }
        
        if (config.GenerateTableOfContents)
        {
            GenerateTableOfContents(document);
        }
        
        if (!string.IsNullOrEmpty(config.WatermarkText))
        {
            AddWatermark(document, config.WatermarkText);
        }
```

执行各种配置的处理步骤。

```csharp
        // 生成输出文件路径
        var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
        var outputFilePath = Path.Combine(outputDirectory, $"{fileName}_processed.docx");
        
        // 保存处理后的Word文档
        document.SaveAs2(outputFilePath);
        Console.WriteLine("  - 已保存处理后的Word文档");
        
        // 如果需要转换为PDF
        if (config.ConvertToPdf)
        {
            var pdfOutputPath = Path.Combine(outputDirectory, $"{fileName}_processed.pdf");
            document.SaveAs2(pdfOutputPath, WdSaveFormat.wdFormatPDF);
            Console.WriteLine("  - 已转换为PDF格式");
        }
    }
```

保存处理结果并根据需要转换为PDF。

```csharp
    private static void StandardizeDocumentFormat(var document)
    {
        var range = document.Range();
        range.Font.Name = "宋体";
        range.Font.Size = 12;
        
        foreach (var paragraph in document.Paragraphs)
        {
            paragraph.Format.LineSpacing = 1.5f;
            paragraph.Format.SpaceAfter = 12;
        }
        
        Console.WriteLine("  - 已标准化文档格式");
    }
```

标准化文档格式。

```csharp
    private static void UpdateDocumentFields(var document)
    {
        document.Range().Fields.Update();
        Console.WriteLine("  - 已更新文档字段");
    }
    
    private static void AddHeaderFooter(var document)
    {
        var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        headerRange.Text = "公司文档";
        headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
```

添加页眉页脚。

```csharp
        var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
        footerRange.Text = " 第 页";
        footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        
        Console.WriteLine("  - 已添加页眉页脚");
    }
```

```csharp
    private static void GenerateTableOfContents(var document)
    {
        // 查找目录位置（假设在文档开头）
        var range = document.Range(0, 0);
        range.Text = "目录\n";
        range.Font.Bold = 1;
        range.Font.Size = 16;
        range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
```

生成目录。

```csharp
        // 插入目录
        document.TablesOfContents.Add(range);
        
        Console.WriteLine("  - 已生成目录");
    }
    
    private static void AddWatermark(var document, string watermarkText)
    {
        // 添加水印（简化实现）
        var watermarkRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
        watermarkRange.Text = watermarkText;
        watermarkRange.Font.Size = 48;
        watermarkRange.Font.Color = WdColor.wdColorGray25;
        watermarkRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        
        Console.WriteLine($"  - 已添加水印: {watermarkText}");
    }
}
```

添加水印。

## 实际应用示例

以下示例演示了完整的文档自动化处理系统：

```csharp
class DocumentAutomationSystem
{
    public static void RunDocumentAutomationDemo()
    {
        Console.WriteLine("=== 文档自动化处理系统演示 ===");
        Console.WriteLine();
        
        try
        {
            // 创建示例文档用于处理
            CreateSampleDocuments();
```

运行文档自动化处理系统演示。

```csharp
            // 1. 批量文档处理演示
            Console.WriteLine("1. 批量文档处理演示");
            BatchDocumentProcessor.ProcessDocumentsInBatch(
                @"C:\temp\SampleDocs\Input",
                @"C:\temp\SampleDocs\Processed"
            );
            Console.WriteLine();
            
            // 2. 文档转换演示
            Console.WriteLine("2. 文档转换演示");
            DocumentConverter.ConvertToPdf(
                @"C:\temp\SampleDocs\Processed",
                @"C:\temp\SampleDocs\PDF"
            );
            Console.WriteLine();
```

执行各种处理演示。

```csharp
            // 3. 自动化工作流演示
            Console.WriteLine("3. 自动化工作流演示");
            var workflowConfig = new DocumentAutomationWorkflow.WorkflowConfiguration
            {
                StandardizeFormat = true,
                UpdateFields = true,
                AddHeaderFooter = true,
                GenerateTableOfContents = true,
                ConvertToPdf = true,
                WatermarkText = "机密文档"
            };
            
            DocumentAutomationWorkflow.ExecuteWorkflow(
                @"C:\temp\SampleDocs\Input",
                @"C:\temp\SampleDocs\WorkflowOutput",
                workflowConfig
            );
            Console.WriteLine();
```

执行自动化工作流演示。

```csharp
            Console.WriteLine("文档自动化处理系统演示完成！");
            Console.WriteLine("处理后的文档位于 C:\\temp\\SampleDocs 目录下");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"系统演示过程中出错: {ex.Message}");
        }
    }
```

```csharp
    private static void CreateSampleDocuments()
    {
        string inputDir = @"C:\temp\SampleDocs\Input";
        
        // 确保目录存在
        if (!Directory.Exists(inputDir))
        {
            Directory.CreateDirectory(inputDir);
        }
        
        // 创建示例文档1
        using (var app = WordFactory.BlankDocument())
        {
            var document = app.ActiveDocument;
            
            document.Range().Text = "示例文档1\n\n这是第一个示例文档的内容。\n包含一些文本用于演示自动化处理。";
```

创建示例文档。

```csharp
            var titleRange = document.Range(0, 6);
            titleRange.Font.Size = 16;
            titleRange.Font.Bold = 1;
            
            document.SaveAs2(Path.Combine(inputDir, "SampleDoc1.docx"));
        }
        
        // 创建示例文档2
        using (var app = WordFactory.BlankDocument())
        {
            var document = app.ActiveDocument;
            
            document.Range().Text = "示例文档2\n\n这是第二个示例文档的内容。\n也包含一些文本用于演示自动化处理。";
            
            var titleRange = document.Range(0, 6);
            titleRange.Font.Size = 16;
            titleRange.Font.Bold = 1;
            
            document.SaveAs2(Path.Combine(inputDir, "SampleDoc2.docx"));
        }
```

创建多个示例文档。

```csharp
        Console.WriteLine("示例文档已创建");
    }
    
    public static void ShowAutomationBestPractices()
    {
        Console.WriteLine("=== 文档自动化处理最佳实践 ===");
        Console.WriteLine();
        
        Console.WriteLine("1. 错误处理:");
        Console.WriteLine("   - 实现全面的异常处理机制");
        Console.WriteLine("   - 记录详细的处理日志");
        Console.WriteLine("   - 支持处理中断后恢复");
        Console.WriteLine();
        
        Console.WriteLine("2. 性能优化:");
        Console.WriteLine("   - 合理使用COM对象生命周期管理");
        Console.WriteLine("   - 批量处理时复用Word实例");
        Console.WriteLine("   - 避免不必要的文档保存操作");
        Console.WriteLine();
```

展示自动化处理最佳实践。

```csharp
        Console.WriteLine("3. 资源管理:");
        Console.WriteLine("   - 正确释放COM对象资源");
        Console.WriteLine("   - 控制并发处理数量");
        Console.WriteLine("   - 监控内存使用情况");
        Console.WriteLine();
        
        Console.WriteLine("4. 质量保证:");
        Console.WriteLine("   - 实现处理前后的文档验证");
        Console.WriteLine("   - 支持处理结果的回滚机制");
        Console.WriteLine("   - 提供处理进度反馈");
    }
}
```

## 应用场景

1. **企业文档管理**：批量处理合同、报告等企业文档
2. **教育机构**：自动化处理试卷、成绩单等教育文档
3. **政府部门**：批量生成通知、公告等政务文档
4. **律师事务所**：标准化处理法律文书和合同模板

## 要点总结

- 批量文档处理可以显著提高工作效率
- 文档转换功能支持多种输出格式
- 自动化工作流可以组合多个处理步骤
- 系统应具备良好的错误处理和日志记录能力
- 需要注意COM资源管理和性能优化

掌握文档自动化处理技能对于构建高效办公系统非常重要，这些功能使开发者能够创建强大的文档处理解决方案。