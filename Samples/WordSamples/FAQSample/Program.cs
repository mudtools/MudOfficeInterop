//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;

namespace FAQSample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("MudTools.OfficeInterop.Word - 常见问题解答示例");

            // 示例1: 一般问题和安装配置问题
            Console.WriteLine("\n=== 示例1: 一般问题和安装配置问题 ===");
            GeneralAndInstallationIssuesDemo();

            // 示例2: 使用问题
            Console.WriteLine("\n=== 示例2: 使用问题 ===");
            UsageIssuesDemo();

            // 示例3: 错误处理和异常管理
            Console.WriteLine("\n=== 示例3: 错误处理和异常管理 ===");
            ErrorHandlingDemo();

            // 示例4: 性能优化建议
            Console.WriteLine("\n=== 示例4: 性能优化建议 ===");
            PerformanceOptimizationDemo();

            // 示例5: API参考信息
            Console.WriteLine("\n=== 示例5: API参考信息 ===");
            ApiReferenceDemo();

            // 示例6: 版本更新和兼容性
            Console.WriteLine("\n=== 示例6: 版本更新和兼容性 ===");
            CompatibilityDemo();

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 一般问题和安装配置问题示例
        /// </summary>
        static void GeneralAndInstallationIssuesDemo()
        {
            try
            {
                // 检查Office安装状态
                Console.WriteLine("1. 检查Office安装状态");
                bool isOfficeInstalled = CommonIssuesHelper.IsOfficeInstalled();
                Console.WriteLine($"Office安装状态: {(isOfficeInstalled ? "已安装" : "未安装")}");

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "FAQSample");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 正确的项目配置示例
                Console.WriteLine("\n2. 正确的项目配置示例");
                Console.WriteLine("项目文件中正确的框架配置示例:");
                Console.WriteLine("<Project Sdk=\"Microsoft.NET.Sdk\">");
                Console.WriteLine("  <PropertyGroup>");
                Console.WriteLine("    <TargetFramework>net6.0-windows</TargetFramework>");
                Console.WriteLine("  </PropertyGroup>");
                Console.WriteLine("</Project>");

                // 正确的资源释放方式
                Console.WriteLine("\n3. 正确的资源释放方式");
                Console.WriteLine("使用using语句或try-finally块确保COM对象正确释放:");

                // 演示正确的资源释放
                using (var app = WordFactory.BlankWorkbook())
                {
                    Console.WriteLine("  - 使用using语句自动释放Word应用程序");
                }

                var app2 = WordFactory.BlankWorkbook();
                try
                {
                    Console.WriteLine("  - 使用try-finally手动释放Word应用程序");
                }
                finally
                {
                    app2.Dispose();
                }

                Console.WriteLine("\n一般问题和安装配置问题示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"一般问题和安装配置问题示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 使用问题示例
        /// </summary>
        static void UsageIssuesDemo()
        {
            try
            {
                // 控制Word窗口可见性
                Console.WriteLine("1. 控制Word窗口可见性");
                using var app = WordFactory.BlankWorkbook();
                Console.WriteLine($"默认可见性: {app.Visible}");
                app.Visible = true; // 显示Word窗口
                Console.WriteLine("已设置Word窗口为可见");

                // 在Web应用中使用STA线程
                Console.WriteLine("\n2. 在Web应用中使用STA线程");
                Console.WriteLine("在STA线程中执行Word操作以避免RPC错误:");
                Console.WriteLine("var task = Task.Factory.StartNew(() => {");
                Console.WriteLine("    Thread.CurrentThread.SetApartmentState(ApartmentState.STA);");
                Console.WriteLine("    // Word操作代码");
                Console.WriteLine("}, TaskCreationOptions.LongRunning);");

                // 性能优化
                Console.WriteLine("\n3. 性能优化");
                app.Visible = false; // 隐藏应用程序界面
                app.ScreenUpdating = false; // 禁用屏幕更新
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone; // 禁用警告

                Console.WriteLine("已应用性能优化设置:");
                Console.WriteLine("  - 隐藏应用程序界面");
                Console.WriteLine("  - 禁用屏幕更新");
                Console.WriteLine("  - 禁用警告");

                app.ScreenUpdating = true; // 恢复屏幕更新

                Console.WriteLine("\n使用问题示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"使用问题示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 错误处理和异常管理示例
        /// </summary>
        static void ErrorHandlingDemo()
        {
            try
            {
                // 处理文件不存在的异常
                Console.WriteLine("1. 处理文件不存在的异常");
                string nonExistentFile = @"C:\NonExistentFile.docx";
                var result = CommonIssuesHelper.HandleFileNotFound(nonExistentFile);
                Console.WriteLine($"文件操作结果: {result.Message} (错误代码: {result.ErrorCode})");

                // 验证文档有效性
                Console.WriteLine("\n2. 验证文档有效性");

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "FAQSample");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建有效文档
                string validDocumentPath = Path.Combine(tempDirectory, "ValidDocument.docx");
                using (var app = WordFactory.BlankWorkbook())
                {
                    app.ActiveDocument.SaveAs(validDocumentPath);
                }

                bool isValid = CommonIssuesHelper.IsDocumentValid(validDocumentPath);
                Console.WriteLine($"文档有效性检查: {validDocumentPath} - {(isValid ? "有效" : "无效")}");

                // 创建无效文档（损坏的文件）
                string invalidDocumentPath = Path.Combine(tempDirectory, "InvalidDocument.docx");
                File.WriteAllText(invalidDocumentPath, "这不是一个有效的Word文档");

                bool isInvalid = CommonIssuesHelper.IsDocumentValid(invalidDocumentPath);
                Console.WriteLine($"文档有效性检查: {invalidDocumentPath} - {(isInvalid ? "有效" : "无效")}");

                Console.WriteLine("\n错误处理和异常管理示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误处理和异常管理示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 性能优化建议示例
        /// </summary>
        static void PerformanceOptimizationDemo()
        {
            try
            {
                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "FAQSample");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建测试文档
                var testDocuments = new List<string>();
                for (int i = 1; i <= 5; i++)
                {
                    string docPath = Path.Combine(tempDirectory, $"TestDocument{i}.docx");
                    using var app = WordFactory.BlankWorkbook();
                    app.ActiveDocument.Range().Text = $"这是测试文档 {i}";
                    app.ActiveDocument.SaveAs(docPath);
                    testDocuments.Add(docPath);
                }

                // 使用高性能文档处理器
                Console.WriteLine("1. 使用高性能文档处理器");
                using (var processor = new PerformanceOptimizationHelper.HighPerformanceDocumentProcessor())
                {
                    var processResult = processor.ProcessDocuments(testDocuments, document =>
                    {
                        // 模拟处理操作
                        document.Range().Text += "\n已处理";
                    });

                    Console.WriteLine($"处理结果: 成功处理 {processResult.ProcessedDocuments.Count} 个文档");
                }

                // 分批处理大型文档
                Console.WriteLine("\n2. 分批处理大型文档");
                var batchResult = PerformanceOptimizationHelper.ProcessLargeDocumentsAsync(
                    testDocuments,
                    2, // 批处理大小
                    document =>
                    {
                        // 模拟处理操作
                        document.Range().Text += "\n批量处理";
                    }).Result;

                Console.WriteLine($"批量处理结果: 总计 {batchResult.TotalDocuments} 个文档，成功 {batchResult.ProcessedDocuments.Count} 个");

                // 生成性能报告
                Console.WriteLine("\n3. 生成性能报告");
                var startTime = DateTime.Now;
                // 模拟一些处理时间
                Task.Delay(100).Wait();
                var endTime = DateTime.Now;
                var processingTime = endTime - startTime;

                string performanceReport = PerformanceOptimizationHelper.GeneratePerformanceReport(batchResult, processingTime);
                Console.WriteLine("性能报告已生成");
                Console.WriteLine(performanceReport);

                Console.WriteLine("性能优化建议示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"性能优化建议示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// API参考信息示例
        /// </summary>
        static void ApiReferenceDemo()
        {
            try
            {
                // WordFactory方法比较
                Console.WriteLine("1. WordFactory方法比较");
                var comparisons = ApiReferenceHelper.GetWordFactoryMethodComparisons();
                foreach (var comparison in comparisons)
                {
                    Console.WriteLine($"方法: {comparison.MethodName}");
                    Console.WriteLine($"  用途: {comparison.Purpose}");
                    Console.WriteLine($"  特点: {comparison.Characteristics}");
                    Console.WriteLine($"  适用场景: {comparison.UseCase}");
                    Console.WriteLine();
                }

                // 遍历文档段落
                Console.WriteLine("2. 遍历文档段落");

                // 创建临时目录
                string tempDirectory = Path.Combine(Path.GetTempPath(), "FAQSample");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                // 创建演示文档
                string demoDocumentPath = Path.Combine(tempDirectory, "DemoDocument.docx");
                bool created = ApiReferenceHelper.CreateDemoDocument(demoDocumentPath);
                Console.WriteLine($"演示文档创建结果: {(created ? "成功" : "失败")}");

                if (created)
                {
                    // 使用for循环遍历段落
                    var paragraphs1 = ApiReferenceHelper.TraverseParagraphs(demoDocumentPath);
                    Console.WriteLine("使用for循环遍历段落:");
                    foreach (var paragraph in paragraphs1)
                    {
                        Console.WriteLine($"  {paragraph}");
                    }

                    // 使用foreach遍历段落
                    var paragraphs2 = ApiReferenceHelper.TraverseParagraphsWithForeach(demoDocumentPath);
                    Console.WriteLine("\n使用foreach遍历段落:");
                    foreach (var paragraph in paragraphs2)
                    {
                        Console.WriteLine($"  {paragraph}");
                    }
                }

                // 生成API使用示例报告
                Console.WriteLine("\n3. 生成API使用示例报告");
                string apiReport = ApiReferenceHelper.GenerateApiUsageReport();
                Console.WriteLine("API使用示例报告已生成");
                Console.WriteLine(apiReport);

                Console.WriteLine("API参考信息示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"API参考信息示例演示出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 版本更新和兼容性示例
        /// </summary>
        static void CompatibilityDemo()
        {
            try
            {
                // Office版本兼容性检查
                Console.WriteLine("1. Office版本兼容性检查");
                var compatibilityResult = CompatibilityHelper.CheckOfficeVersionCompatibility();
                Console.WriteLine($"兼容性检查结果: {compatibilityResult.Message}");

                if (compatibilityResult.VersionInfo != null)
                {
                    Console.WriteLine($"检测到版本: {compatibilityResult.VersionInfo.Name} ({compatibilityResult.VersionInfo.Version})");
                    Console.WriteLine($"发布年份: {compatibilityResult.VersionInfo.ReleaseYear}");
                    Console.WriteLine($"是否支持: {compatibilityResult.VersionInfo.IsSupported}");
                }

                // 支持的Office版本信息
                Console.WriteLine("\n2. 支持的Office版本信息");
                var versions = CompatibilityHelper.GetSupportedOfficeVersions();
                foreach (var version in versions)
                {
                    Console.WriteLine($"  - {version.Name} ({version.Version}) - {(version.IsSupported ? "支持" : "不支持")}");
                }

                // 多语言支持
                Console.WriteLine("\n3. 多语言支持");
                var multiLanguageHelper = new CompatibilityHelper.MultiLanguageSupportHelper();
                Console.WriteLine("多语言支持建议:");
                Console.WriteLine("  - 使用程序化方式而不是UI操作");
                Console.WriteLine("  - 避免依赖特定语言的菜单项或对话框");
                Console.WriteLine("  - 使用常量而不是硬编码的字符串");

                // 功能测试
                Console.WriteLine("\n4. 功能测试");
                var featureTestResult = CompatibilityHelper.TestOfficeFeatures();
                Console.WriteLine($"功能测试结果: {featureTestResult.Message}");

                if (featureTestResult.Success)
                {
                    Console.WriteLine("基本功能测试:");
                    foreach (var feature in featureTestResult.BasicFeatures)
                    {
                        Console.WriteLine($"  {feature.Key}: {(feature.Value ? "支持" : "不支持")}");
                    }

                    Console.WriteLine("高级功能测试:");
                    foreach (var feature in featureTestResult.AdvancedFeatures)
                    {
                        Console.WriteLine($"  {feature.Key}: {(feature.Value ? "支持" : "不支持")}");
                    }
                }

                // 生成兼容性报告
                Console.WriteLine("\n5. 生成兼容性报告");
                string compatibilityReport = CompatibilityHelper.GenerateCompatibilityReport();
                Console.WriteLine("兼容性报告已生成");
                Console.WriteLine(compatibilityReport);

                Console.WriteLine("版本更新和兼容性示例演示完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"版本更新和兼容性示例演示出错: {ex.Message}");
            }
        }
    }
}