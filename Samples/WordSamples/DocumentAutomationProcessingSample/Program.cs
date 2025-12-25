//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace DocumentAutomationProcessingSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 文档自动化处理示例");

            // 示例1: 批量文档处理
            Console.WriteLine("\n=== 示例1: 批量文档处理 ===");
            BatchDocumentProcessingDemo();

            // 示例2: 文档转换
            Console.WriteLine("\n=== 示例2: 文档转换 ===");
            DocumentConversionDemo();

            // 示例3: 自动化工作流
            Console.WriteLine("\n=== 示例3: 自动化工作流 ===");
            AutomationWorkflowDemo();

            // 示例4: 文档内容处理
            Console.WriteLine("\n=== 示例4: 文档内容处理 ===");
            DocumentContentProcessingDemo();

            // 示例5: 实际应用示例
            Console.WriteLine("\n=== 示例5: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例6: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例6: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 批量文档处理示例
        /// </summary>
        static void BatchDocumentProcessingDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DocumentAutomationProcessing");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建示例文档
                string sampleDirectory = Path.Combine(tempDirectory, "Samples");
                if (!Directory.Exists(sampleDirectory))
                {
                    Directory.CreateDirectory(sampleDirectory);

                    var documentInfos = new List<DocumentInfo>
                    {
                        new DocumentInfo {
                            FileName = "文档1.docx",
                            Title = "示例文档1",
                            Content = "这是第一个示例文档的内容。\n包含多行文本。\n用于演示批量处理功能。"
                        },
                        new DocumentInfo {
                            FileName = "文档2.docx",
                            Title = "示例文档2",
                            Content = "这是第二个示例文档的内容。\n也包含多行文本。\n用于演示批量处理功能。"
                        },
                        new DocumentInfo {
                            FileName = "文档3.docx",
                            Title = "示例文档3",
                            Content = "这是第三个示例文档的内容。\n同样包含多行文本。\n用于演示批量处理功能。"
                        }
                    };

                    var creationResult = DocumentProcessor.CreateBatchSampleDocuments(sampleDirectory, documentInfos);
                    Console.WriteLine($"创建示例文档结果: {creationResult.Success}");
                }

                // 执行批量处理
                string processedDirectory = Path.Combine(tempDirectory, "Processed");
                var batchResult = BatchDocumentProcessor.ProcessDocumentsInBatch(
                    sampleDirectory,
                    processedDirectory,
                    "*.docx");

                Console.WriteLine($"批量处理结果: {batchResult.Success}");

                // 生成处理报告
                string report = BatchDocumentProcessor.GenerateProcessingReport(batchResult);
                Console.WriteLine(report);

                // 按类型处理文档
                string typeProcessedDirectory = Path.Combine(tempDirectory, "TypeProcessed");
                var typeResult = BatchDocumentProcessor.ProcessDocumentsByType(
                    sampleDirectory,
                    typeProcessedDirectory,
                    DocumentType.Report);

                Console.WriteLine($"按类型处理结果: {typeResult.Success}");

                Console.WriteLine("批量文档处理示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量文档处理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文档转换示例
        /// </summary>
        static void DocumentConversionDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DocumentAutomationProcessing");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 确保有示例文档
                string sampleDirectory = Path.Combine(tempDirectory, "Samples");
                if (!Directory.Exists(sampleDirectory) || Directory.GetFiles(sampleDirectory, "*.docx").Length == 0)
                {
                    Directory.CreateDirectory(sampleDirectory);

                    var documentInfos = new List<DocumentInfo>
                    {
                        new DocumentInfo {
                            FileName = "转换文档1.docx",
                            Title = "转换文档1",
                            Content = "这是用于转换的第一个示例文档。"
                        },
                        new DocumentInfo {
                            FileName = "转换文档2.docx",
                            Title = "转换文档2",
                            Content = "这是用于转换的第二个示例文档。"
                        }
                    };

                    DocumentProcessor.CreateBatchSampleDocuments(sampleDirectory, documentInfos);
                }

                // 转换为PDF
                string pdfDirectory = Path.Combine(tempDirectory, "PDF");
                var pdfResult = DocumentConverter.ConvertToPdf(sampleDirectory, pdfDirectory);
                Console.WriteLine($"转换为PDF结果: {pdfResult.Success}");

                // 转换为HTML
                string htmlDirectory = Path.Combine(tempDirectory, "HTML");
                var htmlResult = DocumentConverter.ConvertToHtml(sampleDirectory, htmlDirectory);
                Console.WriteLine($"转换为HTML结果: {htmlResult.Success}");

                // 转换为多种格式
                string multiFormatDirectory = Path.Combine(tempDirectory, "MultiFormat");
                var formats = new List<WdSaveFormat>
                {
                    WdSaveFormat.wdFormatPDF,
                    WdSaveFormat.wdFormatRTF,
                    WdSaveFormat.wdFormatFilteredHTML
                };

                var multiFormatResult = DocumentConverter.ConvertToMultipleFormats(
                    sampleDirectory,
                    multiFormatDirectory,
                    formats);

                Console.WriteLine($"多格式转换结果: {multiFormatResult.Success}");

                // 生成转换报告
                string pdfReport = DocumentConverter.GenerateConversionReport(pdfResult);
                Console.WriteLine(pdfReport);

                string multiFormatReport = DocumentConverter.GenerateMultiFormatConversionReport(multiFormatResult);
                Console.WriteLine(multiFormatReport);

                Console.WriteLine("文档转换示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档转换示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 自动化工作流示例
        /// </summary>
        static void AutomationWorkflowDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DocumentAutomationProcessing");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 确保有示例文档
                string sampleDirectory = Path.Combine(tempDirectory, "Samples");
                if (!Directory.Exists(sampleDirectory) || Directory.GetFiles(sampleDirectory, "*.docx").Length == 0)
                {
                    Directory.CreateDirectory(sampleDirectory);

                    var documentInfos = new List<DocumentInfo>
                    {
                        new DocumentInfo {
                            FileName = "工作流文档1.docx",
                            Title = "工作流文档1",
                            Content = "这是用于工作流处理的第一个示例文档。"
                        },
                        new DocumentInfo {
                            FileName = "工作流文档2.docx",
                            Title = "工作流文档2",
                            Content = "这是用于工作流处理的第二个示例文档。"
                        }
                    };

                    DocumentProcessor.CreateBatchSampleDocuments(sampleDirectory, documentInfos);
                }

                // 执行标准工作流
                var standardConfig = new DocumentAutomationWorkflow.WorkflowConfiguration
                {
                    StandardizeFormat = true,
                    UpdateFields = true,
                    AddHeaderFooter = true,
                    ConvertToPdf = true,
                    GenerateTableOfContents = false,
                    WatermarkText = null
                };

                string standardWorkflowDirectory = Path.Combine(tempDirectory, "StandardWorkflow");
                var standardResult = DocumentAutomationWorkflow.ExecuteWorkflow(
                    sampleDirectory,
                    standardWorkflowDirectory,
                    standardConfig);

                Console.WriteLine($"标准工作流执行结果: {standardResult.Success}");

                // 执行带水印的工作流
                var watermarkConfig = new DocumentAutomationWorkflow.WorkflowConfiguration
                {
                    StandardizeFormat = true,
                    UpdateFields = true,
                    AddHeaderFooter = true,
                    ConvertToPdf = false,
                    GenerateTableOfContents = true,
                    WatermarkText = "机密"
                };

                string watermarkWorkflowDirectory = Path.Combine(tempDirectory, "WatermarkWorkflow");
                var watermarkResult = DocumentAutomationWorkflow.ExecuteWorkflow(
                    sampleDirectory,
                    watermarkWorkflowDirectory,
                    watermarkConfig);

                Console.WriteLine($"带水印工作流执行结果: {watermarkResult.Success}");

                // 创建并执行自定义工作流
                var customSteps = new List<WorkflowStep>
                {
                    new WorkflowStep { Type = WorkflowStepType.StandardizeFormat },
                    new WorkflowStep { Type = WorkflowStepType.AddHeaderFooter },
                    new WorkflowStep { Type = WorkflowStepType.AddWatermark, Parameter = "草稿" }
                };

                var customWorkflow = DocumentAutomationWorkflow.CreateCustomWorkflow("自定义工作流", customSteps);

                string customWorkflowDirectory = Path.Combine(tempDirectory, "CustomWorkflow");
                var customResult = DocumentAutomationWorkflow.ExecuteCustomWorkflow(
                    sampleDirectory,
                    customWorkflowDirectory,
                    customWorkflow);

                Console.WriteLine($"自定义工作流执行结果: {customResult.Success}");

                // 生成工作流报告
                string standardReport = DocumentAutomationWorkflow.GenerateWorkflowReport(standardResult);
                Console.WriteLine(standardReport);

                Console.WriteLine("自动化工作流示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自动化工作流示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 文档内容处理示例
        /// </summary>
        static void DocumentContentProcessingDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DocumentAutomationProcessing");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建示例文档
                string sampleFilePath = Path.Combine(tempDirectory, "内容处理示例.docx");
                DocumentProcessor.CreateSampleDocument(
                    sampleFilePath,
                    "内容处理示例文档",
                    "这是一个用于内容处理的示例文档。\n文档包含多行文本。\n可以进行替换、插入、删除等操作。");

                // 打开文档进行内容处理
                using var app = WordFactory.Open(sampleFilePath);
                using var document = app.ActiveDocument;

                // 定义内容操作
                var operations = new List<ContentOperation>
                {
                    new ContentOperation {
                        Type = ContentOperationType.ReplaceText,
                        Parameters = new Dictionary<string, string>
                        {
                            { "FindText", "内容处理" },
                            { "ReplaceText", "内容编辑" }
                        },
                        Description = "替换文本: 内容处理 -> 内容编辑"
                    },
                    new ContentOperation {
                        Type = ContentOperationType.InsertText,
                        Parameters = new Dictionary<string, string>
                        {
                            { "Position", document.Content.End.ToString() },
                            { "Text", "\n这是插入的新文本。" }
                        },
                        Description = "在文档末尾插入文本"
                    },
                    new ContentOperation {
                        Type = ContentOperationType.FormatText,
                        Parameters = new Dictionary<string, string>
                        {
                            { "Start", "0" },
                            { "End", "10" },
                            { "FontName", "微软雅黑" },
                            { "FontSize", "14" },
                            { "Bold", "true" }
                        },
                        Description = "格式化文档开头的文本"
                    }
                };

                // 执行内容处理
                var processingResult = DocumentProcessor.ProcessDocumentContent(document, operations);
                Console.WriteLine($"内容处理结果: {processingResult.Success}");

                // 保存处理后的文档
                string processedFilePath = Path.Combine(tempDirectory, "内容处理后.docx");
                document.SaveAs(processedFilePath);


                Console.WriteLine("文档内容处理示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"文档内容处理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                Console.WriteLine("=== 文档自动化处理系统演示 ===");
                Console.WriteLine();

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DocumentAutomationProcessing");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 步骤1: 创建示例文档
                Console.WriteLine("步骤1: 创建示例文档");
                string sampleDirectory = Path.Combine(tempDirectory, "Samples");
                var documentInfos = new List<DocumentInfo>
                {
                    new DocumentInfo {
                        FileName = "月度报告1.docx",
                        Title = "2023年1月月度报告",
                        Content = "这是2023年1月的月度报告内容。\n报告包含各种业务数据和分析。"
                    },
                    new DocumentInfo {
                        FileName = "月度报告2.docx",
                        Title = "2023年2月月度报告",
                        Content = "这是2023年2月的月度报告内容。\n报告包含各种业务数据和分析。"
                    },
                    new DocumentInfo {
                        FileName = "月度报告3.docx",
                        Title = "2023年3月月度报告",
                        Content = "这是2023年3月的月度报告内容。\n报告包含各种业务数据和分析。"
                    }
                };

                DocumentProcessor.CreateBatchSampleDocuments(sampleDirectory, documentInfos);
                Console.WriteLine();

                // 步骤2: 批量处理文档
                Console.WriteLine("步骤2: 批量处理文档");
                string processedDirectory = Path.Combine(tempDirectory, "Processed");
                var batchResult = BatchDocumentProcessor.ProcessDocumentsInBatch(
                    sampleDirectory,
                    processedDirectory,
                    "*.docx");

                string batchReport = BatchDocumentProcessor.GenerateProcessingReport(batchResult);
                Console.WriteLine(batchReport);
                Console.WriteLine();

                // 步骤3: 转换文档格式
                Console.WriteLine("步骤3: 转换文档格式");
                string pdfDirectory = Path.Combine(tempDirectory, "PDF");
                var pdfResult = DocumentConverter.ConvertToPdf(processedDirectory, pdfDirectory);

                string pdfReport = DocumentConverter.GenerateConversionReport(pdfResult);
                Console.WriteLine(pdfReport);
                Console.WriteLine();

                // 步骤4: 执行自动化工作流
                Console.WriteLine("步骤4: 执行自动化工作流");
                var workflowConfig = new DocumentAutomationWorkflow.WorkflowConfiguration
                {
                    StandardizeFormat = true,
                    UpdateFields = true,
                    AddHeaderFooter = true,
                    ConvertToPdf = true,
                    GenerateTableOfContents = false,
                    WatermarkText = "内部使用"
                };

                string workflowDirectory = Path.Combine(tempDirectory, "Workflow");
                var workflowResult = DocumentAutomationWorkflow.ExecuteWorkflow(
                    processedDirectory,
                    workflowDirectory,
                    workflowConfig);

                string workflowReport = DocumentAutomationWorkflow.GenerateWorkflowReport(workflowResult);
                Console.WriteLine(workflowReport);
                Console.WriteLine();

                Console.WriteLine("文档自动化处理系统演示完成！");
                Console.WriteLine($"生成的文档位于 {tempDirectory} 目录下");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"实际应用示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用辅助类的完整示例
        /// </summary>
        static void CompleteExampleWithHelpers()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "DocumentAutomationProcessing");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 1. 创建示例文档
                Console.WriteLine("1. 创建示例文档");
                string sampleDirectory = Path.Combine(tempDirectory, "CompleteSamples");
                var documentInfos = new List<DocumentInfo>
                {
                    new DocumentInfo {
                        FileName = "完整示例1.docx",
                        Title = "完整示例文档1",
                        Content = "这是完整示例的第一个文档。\n用于演示所有功能。"
                    },
                    new DocumentInfo {
                        FileName = "完整示例2.docx",
                        Title = "完整示例文档2",
                        Content = "这是完整示例的第二个文档。\n用于演示所有功能。"
                    },
                    new DocumentInfo {
                        FileName = "完整示例3.docx",
                        Title = "完整示例文档3",
                        Content = "这是完整示例的第三个文档。\n用于演示所有功能。"
                    }
                };

                var creationResult = DocumentProcessor.CreateBatchSampleDocuments(sampleDirectory, documentInfos);
                Console.WriteLine($"创建文档结果: 成功创建 {creationResult.CreatedDocuments.Count} 个文档");
                Console.WriteLine();

                // 2. 批量处理文档
                Console.WriteLine("2. 批量处理文档");
                string processedDirectory = Path.Combine(tempDirectory, "CompleteProcessed");
                var batchResult = BatchDocumentProcessor.ProcessDocumentsInBatch(
                    sampleDirectory,
                    processedDirectory,
                    "*.docx");

                Console.WriteLine($"批量处理结果: 处理了 {batchResult.ProcessedFiles.Count} 个文档");
                Console.WriteLine();

                // 3. 转换文档格式
                Console.WriteLine("3. 转换文档格式");
                string convertedDirectory = Path.Combine(tempDirectory, "CompleteConverted");
                var formats = new List<WdSaveFormat>
                {
                    WdSaveFormat.wdFormatPDF,
                    WdSaveFormat.wdFormatRTF
                };

                var conversionResult = DocumentConverter.ConvertToMultipleFormats(
                    processedDirectory,
                    convertedDirectory,
                    formats);

                Console.WriteLine($"多格式转换结果: 转换了 {conversionResult.FormatResults.Count} 种格式");
                Console.WriteLine();

                // 4. 执行自动化工作流
                Console.WriteLine("4. 执行自动化工作流");
                var workflowConfig = new DocumentAutomationWorkflow.WorkflowConfiguration
                {
                    StandardizeFormat = true,
                    UpdateFields = true,
                    AddHeaderFooter = true,
                    ConvertToPdf = false,
                    GenerateTableOfContents = true,
                    WatermarkText = "示例"
                };

                string workflowDirectory = Path.Combine(tempDirectory, "CompleteWorkflow");
                var workflowResult = DocumentAutomationWorkflow.ExecuteWorkflow(
                    processedDirectory,
                    workflowDirectory,
                    workflowConfig);

                Console.WriteLine($"工作流执行结果: 处理了 {workflowResult.ProcessedFiles.Count} 个文档");
                Console.WriteLine();

                // 5. 内容处理示例
                Console.WriteLine("5. 内容处理示例");
                string sampleFilePath = Path.Combine(processedDirectory, "完整示例1_processed.docx");
                if (File.Exists(sampleFilePath))
                {
                    using var app = WordFactory.Open(sampleFilePath);
                    using var document = app.ActiveDocument;

                    var operations = new List<ContentOperation>
                    {
                        new ContentOperation {
                            Type = ContentOperationType.ReplaceText,
                            Parameters = new Dictionary<string, string>
                            {
                                { "FindText", "完整示例" },
                                { "ReplaceText", "完整处理" }
                            },
                            Description = "替换文本"
                        }
                    };

                    var processingResult = DocumentProcessor.ProcessDocumentContent(document, operations);
                    Console.WriteLine($"内容处理结果: 执行了 {processingResult.CompletedOperations.Count} 个操作");

                }
                Console.WriteLine();

                // 6. 生成所有报告
                Console.WriteLine("6. 生成所有报告");
                string batchReport = BatchDocumentProcessor.GenerateProcessingReport(batchResult);
                string conversionReport = DocumentConverter.GenerateMultiFormatConversionReport(conversionResult);
                string workflowReport = DocumentAutomationWorkflow.GenerateWorkflowReport(workflowResult);

                Console.WriteLine("批量处理报告:");
                Console.WriteLine(batchReport);
                Console.WriteLine();

                Console.WriteLine("格式转换报告:");
                Console.WriteLine(conversionReport);
                Console.WriteLine();

                Console.WriteLine("工作流报告:");
                Console.WriteLine(workflowReport);
                Console.WriteLine();

                Console.WriteLine("使用辅助类的完整示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用辅助类的完整示例演示出错: {ex.Message}");
            }
        }
    }
}