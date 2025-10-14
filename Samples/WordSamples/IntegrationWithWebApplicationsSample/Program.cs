//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace IntegrationWithWebApplicationsSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 集成到Web应用示例");

            // 示例1: Word文档服务
            Console.WriteLine("\n=== 示例1: Word文档服务 ===");
            WordDocumentServiceDemo();

            // 示例2: 线程安全处理
            Console.WriteLine("\n=== 示例2: 线程安全处理 ===");
            ThreadSafeProcessingDemo();

            // 示例3: 资源管理
            Console.WriteLine("\n=== 示例3: 资源管理 ===");
            ResourceManagementDemo();

            // 示例4: 实际应用示例
            Console.WriteLine("\n=== 示例4: 实际应用示例 ===");
            RealWorldApplicationDemo();

            // 示例5: 使用辅助类的完整示例
            Console.WriteLine("\n=== 示例5: 使用辅助类的完整示例 ===");
            CompleteExampleWithHelpers();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// Word文档服务示例
        /// </summary>
        static void WordDocumentServiceDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "IntegrationWithWebApplications");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                var documentService = new WordDocumentService();

                // 创建文档
                Console.WriteLine("1. 创建文档");
                string content = "这是一个在Web应用中创建的Word文档。\n\n" +
                                "文档内容可以包含多行文本。\n" +
                                "支持各种格式化选项。";
                string documentPath = documentService.CreateDocumentAsync(content).Result;
                Console.WriteLine($"文档已创建: {documentPath}");

                // 从模板创建文档
                Console.WriteLine("\n2. 从模板创建文档");
                var templateData = new Dictionary<string, string>
                {
                    { "Name", "张三" },
                    { "Position", "软件工程师" },
                    { "Department", "技术部" }
                };
                string templatePath = Path.Combine(tempDirectory, "Template.docx");

                // 创建一个简单的模板文档
                using (var app = WordFactory.BlankWorkbook())
                {
                    var document = app.ActiveDocument;
                    document.Range().Text = "员工信息\n\n" +
                                          "姓名: {Name}\n" +
                                          "职位: {Position}\n" +
                                          "部门: {Department}\n";
                    document.SaveAs(templatePath);
                }

                string templateDocumentPath = documentService.CreateDocumentFromTemplateAsync(templatePath, templateData).Result;
                Console.WriteLine($"从模板创建的文档: {templateDocumentPath}");

                // 转换文档格式
                Console.WriteLine("\n3. 转换文档格式");
                string pdfPath = documentService.ConvertDocumentAsync(documentPath, WdSaveFormat.wdFormatPDF).Result;
                Console.WriteLine($"文档已转换为PDF: {pdfPath}");

                // 批量处理文档
                Console.WriteLine("\n4. 批量处理文档");
                var documents = new List<DocumentInfo>
                {
                    new DocumentInfo { Name = "文档1", Content = "这是第一个文档的内容" },
                    new DocumentInfo { Name = "文档2", Content = "这是第二个文档的内容" },
                    new DocumentInfo { Name = "文档3", Content = "这是第三个文档的内容" }
                };

                var batchResult = documentService.ProcessDocumentsAsync(documents).Result;
                Console.WriteLine($"批量处理结果: 总计{batchResult.TotalDocuments}个文档，成功{batchResult.ProcessedDocuments.Count}个，失败{batchResult.FailedDocuments.Count}个");

                Console.WriteLine("\nWord文档服务示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Word文档服务示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 线程安全处理示例
        /// </summary>
        static void ThreadSafeProcessingDemo()
        {
            try
            {
                var threadSafeService = new ThreadSafeWordService();

                // 在STA线程中执行操作
                Console.WriteLine("1. 在STA线程中执行操作");
                string content = "这是在线程安全环境中创建的文档。\n\n" +
                                "通过STA线程确保Office COM组件的正确使用。";

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "IntegrationWithWebApplications");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建模板
                string templatePath = Path.Combine(tempDirectory, "ThreadSafeTemplate.docx");
                using (var app = WordFactory.BlankWorkbook())
                {
                    var document = app.ActiveDocument;
                    document.Range().Text = "线程安全文档模板\n\n" +
                                          "内容: {Content}\n" +
                                          "创建时间: {CreateTime}\n";
                    document.SaveAs(templatePath);
                }

                var templateData = new Dictionary<string, string>
                {
                    { "Content", "这是从线程安全服务生成的内容" },
                    { "CreateTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
                };

                string documentPath = threadSafeService.ProcessDocumentAsync(templatePath, templateData).Result;
                Console.WriteLine($"线程安全处理的文档: {documentPath}");

                // 批量处理文档
                Console.WriteLine("\n2. 批量处理文档");
                var requests = new List<DocumentProcessingRequest>
                {
                    new DocumentProcessingRequest
                    {
                        DocumentName = "批量文档1",
                        TemplatePath = templatePath,
                        Data = new Dictionary<string, string>
                        {
                            { "Content", "批量处理的第一个文档" },
                            { "CreateTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
                        }
                    },
                    new DocumentProcessingRequest
                    {
                        DocumentName = "批量文档2",
                        TemplatePath = templatePath,
                        Data = new Dictionary<string, string>
                        {
                            { "Content", "批量处理的第二个文档" },
                            { "CreateTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
                        }
                    }
                };

                var results = threadSafeService.ProcessDocumentsAsync(requests).Result;
                foreach (var result in results)
                {
                    Console.WriteLine($"  {result.DocumentName}: {(result.Success ? "成功" : "失败")}");
                    if (!result.Success)
                    {
                        Console.WriteLine($"    错误: {result.ErrorMessage}");
                    }
                }

                Console.WriteLine("\n线程安全处理示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"线程安全处理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 资源管理示例
        /// </summary>
        static void ResourceManagementDemo()
        {
            try
            {
                using (var resourceService = new ResourceManagedWordService())
                {
                    // 生成文档
                    Console.WriteLine("1. 生成文档");
                    string content = "这是一个使用资源管理服务创建的文档。\n\n" +
                                    "服务确保正确释放Word应用程序资源。";
                    string documentPath = resourceService.GenerateDocument(content);
                    Console.WriteLine($"资源管理服务创建的文档: {documentPath}");

                    // 从模板生成文档
                    Console.WriteLine("\n2. 从模板生成文档");

                    // 创建临时目录
                    string tempDirectory = Path.Combine(Path.GetTempPath(), "IntegrationWithWebApplications");
                    if (!Directory.Exists(tempDirectory))
                    {
                        Directory.CreateDirectory(tempDirectory);
                    }

                    // 创建模板
                    string templatePath = Path.Combine(tempDirectory, "ResourceTemplate.docx");
                    using (var app = WordFactory.BlankWorkbook())
                    {
                        var document = app.ActiveDocument;
                        document.Range().Text = "资源管理模板\n\n" +
                                              "姓名: {Name}\n" +
                                              "职位: {Position}\n";
                        document.SaveAs(templatePath);
                    }

                    var templateData = new Dictionary<string, string>
                    {
                        { "Name", "李四" },
                        { "Position", "项目经理" }
                    };

                    string templateDocumentPath = resourceService.GenerateDocumentFromTemplate(templatePath, templateData);
                    Console.WriteLine($"从模板生成的文档: {templateDocumentPath}");

                    // 执行文档操作
                    Console.WriteLine("\n3. 执行文档操作");
                    resourceService.ExecuteDocumentOperation(document =>
                    {
                        document.Range().Text = "通过资源管理服务执行的操作\n\n" +
                                              "这是在操作中添加的内容。";
                        document.Paragraphs[1].Range.Font.Bold = true;
                        document.Paragraphs[1].Range.Font.Size = 16;
                    });
                    Console.WriteLine("文档操作已执行");

                    Console.WriteLine("\n资源管理示例演示完成");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"资源管理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 实际应用示例
        /// </summary>
        static void RealWorldApplicationDemo()
        {
            try
            {
                Console.WriteLine("=== Web应用中文档处理系统演示 ===");
                Console.WriteLine();

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "IntegrationWithWebApplications");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                var logger = new ConsoleLogger();
                var documentGenerationService = new DocumentGenerationService(logger);

                // 步骤1: 创建文档请求
                Console.WriteLine("步骤1: 创建文档请求");
                var documentRequest = new DocumentRequest
                {
                    Title = "项目报告",
                    Author = "王五",
                    Content = "这是项目报告的主要内容。\n\n报告涵盖了项目的各个方面，包括进度、问题和下一步计划。",
                    Sections = new List<DocumentSection>
                    {
                        new DocumentSection
                        {
                            Heading = "项目进度",
                            Text = "项目目前完成了70%，按计划进行。"
                        },
                        new DocumentSection
                        {
                            Heading = "遇到的问题",
                            Text = "遇到了一些技术难题，但已找到解决方案。"
                        },
                        new DocumentSection
                        {
                            Heading = "下一步计划",
                            Text = "继续开发剩余功能，并进行测试。"
                        }
                    },
                    CustomFields = new Dictionary<string, string>
                    {
                        { "ProjectName", "文档处理系统" },
                        { "Version", "1.0" }
                    }
                };

                // 步骤2: 生成文档
                Console.WriteLine("\n步骤2: 生成文档");
                var generationResult = documentGenerationService.GenerateDocumentAsync(documentRequest).Result;
                Console.WriteLine($"文档生成结果: {generationResult.Success}");
                if (generationResult.Success)
                {
                    Console.WriteLine($"文档路径: {generationResult.FilePath}");
                }
                else
                {
                    Console.WriteLine($"错误信息: {generationResult.Message}");
                }

                // 步骤3: 创建模板并从模板生成文档
                Console.WriteLine("\n步骤3: 从模板生成文档");
                string templatePath = Path.Combine(tempDirectory, "WebAppTemplate.docx");
                using (var app = WordFactory.BlankWorkbook())
                {
                    var document = app.ActiveDocument;
                    document.Range().Text = "项目报告\n\n" +
                                          "项目名称: {{ProjectName}}\n" +
                                          "版本: {{Version}}\n" +
                                          "作者: {{Author}}\n" +
                                          "日期: {{Date}}\n\n" +
                                          "{{Content}}\n\n" +
                                          "报告结束";
                    document.SaveAs(templatePath);
                }

                var templateRequest = new DocumentRequest
                {
                    Title = "模板生成的报告",
                    Author = "赵六",
                    Content = "这是通过模板生成的报告内容。\n\n模板中使用了占位符，可以在生成时替换为实际内容。",
                    CustomFields = new Dictionary<string, string>
                    {
                        { "ProjectName", "Web应用集成项目" },
                        { "Version", "2.0" }
                    }
                };

                var templateResult = documentGenerationService.GenerateDocumentFromTemplateAsync(templatePath, templateRequest).Result;
                Console.WriteLine($"模板文档生成结果: {templateResult.Success}");
                if (templateResult.Success)
                {
                    Console.WriteLine($"文档路径: {templateResult.FilePath}");
                }

                // 步骤4: 批量生成文档
                Console.WriteLine("\n步骤4: 批量生成文档");
                var batchRequests = new List<DocumentRequest>
                {
                    new DocumentRequest
                    {
                        Title = "周报1",
                        Author = "员工A",
                        Content = "本周完成了任务1和任务2。"
                    },
                    new DocumentRequest
                    {
                        Title = "周报2",
                        Author = "员工B",
                        Content = "本周完成了任务3和任务4。"
                    },
                    new DocumentRequest
                    {
                        Title = "周报3",
                        Author = "员工C",
                        Content = "本周完成了任务5和任务6。"
                    }
                };

                var batchResults = documentGenerationService.GenerateDocumentsAsync(batchRequests).Result;
                Console.WriteLine($"批量生成完成，总计: {batchResults.Count} 个文档");
                int successCount = batchResults.Count(r => r.Success);
                Console.WriteLine($"成功: {successCount} 个，失败: {batchResults.Count - successCount} 个");

                // 步骤5: 转换文档格式
                Console.WriteLine("\n步骤5: 转换文档格式");
                if (generationResult.Success && File.Exists(generationResult.FilePath))
                {
                    var pdfResult = documentGenerationService.ConvertDocumentAsync(generationResult.FilePath, WdSaveFormat.wdFormatPDF).Result;
                    Console.WriteLine($"PDF转换结果: {pdfResult.Success}");
                    if (pdfResult.Success)
                    {
                        Console.WriteLine($"PDF文件路径: {pdfResult.FilePath}");
                    }
                }

                Console.WriteLine("\nWeb应用中文档处理系统演示完成！");
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
                string tempDirectory = Path.Combine(Path.GetTempPath(), "IntegrationWithWebApplications");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                var logger = new ConsoleLogger();

                // 1. Word文档服务
                Console.WriteLine("1. Word文档服务");
                var documentService = new WordDocumentService();
                string content = "使用辅助类创建的文档\n\n" +
                                "这是通过WordDocumentService创建的文档。";
                string documentPath = documentService.CreateDocumentAsync(content).Result;
                Console.WriteLine($"文档服务创建的文档: {documentPath}");
                Console.WriteLine();

                // 2. 线程安全服务
                Console.WriteLine("2. 线程安全服务");
                var threadSafeService = new ThreadSafeWordService();

                // 创建模板
                string templatePath = Path.Combine(tempDirectory, "CompleteExampleTemplate.docx");
                using (var app = WordFactory.BlankWorkbook())
                {
                    var document = app.ActiveDocument;
                    document.Range().Text = "完整示例模板\n\n" +
                                          "内容: {Content}\n" +
                                          "时间: {Time}\n";
                    document.SaveAs(templatePath);
                }

                var templateData = new Dictionary<string, string>
                {
                    { "Content", "通过线程安全服务处理的内容" },
                    { "Time", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
                };

                string threadSafeDocumentPath = threadSafeService.ProcessDocumentAsync(templatePath, templateData).Result;
                Console.WriteLine($"线程安全服务处理的文档: {threadSafeDocumentPath}");
                Console.WriteLine();

                // 3. 资源管理服务
                Console.WriteLine("3. 资源管理服务");
                using (var resourceService = new ResourceManagedWordService())
                {
                    string resourceDocumentPath = resourceService.GenerateDocument("资源管理服务创建的文档\n\n内容...");
                    Console.WriteLine($"资源管理服务创建的文档: {resourceDocumentPath}");
                }
                Console.WriteLine();

                // 4. 文档生成服务
                Console.WriteLine("4. 文档生成服务");
                var documentGenerationService = new DocumentGenerationService(logger);
                var documentRequest = new DocumentRequest
                {
                    Title = "完整示例报告",
                    Author = "示例用户",
                    Content = "这是使用所有辅助类创建的完整示例报告。"
                };

                var generationResult = documentGenerationService.GenerateDocumentAsync(documentRequest).Result;
                Console.WriteLine($"文档生成服务结果: {generationResult.Success}");
                if (generationResult.Success)
                {
                    Console.WriteLine($"文档路径: {generationResult.FilePath}");
                }
                Console.WriteLine();

                // 5. 批量处理
                Console.WriteLine("5. 批量处理");
                var batchDocuments = new List<DocumentInfo>
                {
                    new DocumentInfo { Name = "批量1", Content = "批量处理的第一个文档" },
                    new DocumentInfo { Name = "批量2", Content = "批量处理的第二个文档" }
                };

                var batchResult = documentService.ProcessDocumentsAsync(batchDocuments).Result;
                Console.WriteLine($"批量处理结果: 成功{batchResult.ProcessedDocuments.Count}个，失败{batchResult.FailedDocuments.Count}个");
                Console.WriteLine();

                // 6. 格式转换
                Console.WriteLine("6. 格式转换");
                if (generationResult.Success && File.Exists(generationResult.FilePath))
                {
                    var pdfResult = documentGenerationService.ConvertDocumentAsync(generationResult.FilePath, WdSaveFormat.wdFormatPDF).Result;
                    Console.WriteLine($"格式转换结果: {pdfResult.Success}");
                    if (pdfResult.Success)
                    {
                        Console.WriteLine($"PDF文件路径: {pdfResult.FilePath}");
                    }
                }
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