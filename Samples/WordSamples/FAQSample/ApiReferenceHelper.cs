//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using System.Text;

namespace FAQSample
{
    /// <summary>
    /// API参考帮助类
    /// </summary>
    public class ApiReferenceHelper
    {
        /// <summary>
        /// WordFactory方法比较信息
        /// </summary>
        public class WordFactoryMethodComparison
        {
            /// <summary>
            /// 方法名称
            /// </summary>
            public string MethodName { get; set; }

            /// <summary>
            /// 用途
            /// </summary>
            public string Purpose { get; set; }

            /// <summary>
            /// 特点
            /// </summary>
            public string Characteristics { get; set; }

            /// <summary>
            /// 适用场景
            /// </summary>
            public string UseCase { get; set; }
        }

        /// <summary>
        /// 获取WordFactory方法比较信息
        /// </summary>
        /// <returns>方法比较信息列表</returns>
        public static List<WordFactoryMethodComparison> GetWordFactoryMethodComparisons()
        {
            return new List<WordFactoryMethodComparison>
            {
                new WordFactoryMethodComparison
                {
                    MethodName = "BlankWorkbook()",
                    Purpose = "创建空白文档",
                    Characteristics = "启动Word并创建新文档",
                    UseCase = "需要从头开始创建新文档时使用"
                },
                new WordFactoryMethodComparison
                {
                    MethodName = "CreateFrom(string)",
                    Purpose = "基于模板创建",
                    Characteristics = "从.dotx模板创建新文档",
                    UseCase = "需要基于特定模板创建文档时使用"
                },
                new WordFactoryMethodComparison
                {
                    MethodName = "Open(string)",
                    Purpose = "打开现有文档",
                    Characteristics = "打开已存在的.docx文件",
                    UseCase = "需要编辑或处理现有文档时使用"
                },
                new WordFactoryMethodComparison
                {
                    MethodName = "Connection(object)",
                    Purpose = "连接现有实例",
                    Characteristics = "连接到已运行的Word实例",
                    UseCase = "需要与已经运行的Word实例交互时使用"
                }
            };
        }

        /// <summary>
        /// 遍历文档段落
        /// </summary>
        /// <param name="documentPath">文档路径</param>
        /// <returns>段落文本列表</returns>
        public static List<string> TraverseParagraphs(string documentPath)
        {
            var paragraphs = new List<string>();

            try
            {
                using var app = WordFactory.Open(documentPath);
                var document = app.ActiveDocument;

                // 方法1：使用for循环
                for (int i = 1; i <= document.Paragraphs.Count; i++)
                {
                    var paragraph = document.Paragraphs[i];
                    paragraphs.Add($"段落 {i}: {paragraph.Range.Text.Trim()}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"遍历段落时出错: {ex.Message}");
            }

            return paragraphs;
        }

        /// <summary>
        /// 使用foreach遍历文档段落
        /// </summary>
        /// <param name="documentPath">文档路径</param>
        /// <returns>段落文本列表</returns>
        public static List<string> TraverseParagraphsWithForeach(string documentPath)
        {
            var paragraphs = new List<string>();

            try
            {
                using var app = WordFactory.Open(documentPath);
                var document = app.ActiveDocument;

                // 方法2：使用foreach（如果支持）
                int index = 1;
                foreach (var paragraph in document.Paragraphs)
                {
                    paragraphs.Add($"段落 {index}: {paragraph.Range.Text.Trim()}");
                    index++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"遍历段落时出错: {ex.Message}");
            }

            return paragraphs;
        }

        /// <summary>
        /// 创建演示文档
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否创建成功</returns>
        public static bool CreateDemoDocument(string filePath)
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加示例内容
                document.Range().Text = "这是第一段文本。\n\n" +
                                      "这是第二段文本，包含一些内容。\n\n" +
                                      "这是第三段文本，也是最后一段。";

                // 格式化第一段
                var firstParagraph = document.Paragraphs[1];
                firstParagraph.Range.Font.Bold = 1;
                firstParagraph.Range.Font.Size = 14;

                // 保存文档
                document.SaveAs2(filePath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建演示文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 比较不同WordFactory方法的性能
        /// </summary>
        /// <param name="iterations">迭代次数</param>
        /// <returns>性能比较结果</returns>
        public static PerformanceComparisonResult CompareWordFactoryMethods(int iterations)
        {
            var result = new PerformanceComparisonResult();

            // 测试BlankWorkbook方法
            var blankWorkbookTimes = new List<TimeSpan>();
            for (int i = 0; i < iterations; i++)
            {
                var startTime = DateTime.Now;
                try
                {
                    using var app = WordFactory.BlankWorkbook();
                    // 简单操作
                    var _ = app.ActiveDocument.Paragraphs.Count;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"BlankWorkbook测试出错: {ex.Message}");
                }
                var endTime = DateTime.Now;
                blankWorkbookTimes.Add(endTime - startTime);
            }

            result.BlankWorkbookAverageTime = blankWorkbookTimes.Average(t => t.TotalMilliseconds);

            // 测试Open方法（需要先创建测试文件）
            var tempPath = Path.GetTempFileName().Replace(".tmp", ".docx");
            try
            {
                // 创建测试文件
                using (var app = WordFactory.BlankWorkbook())
                {
                    app.ActiveDocument.SaveAs2(tempPath);
                }

                var openTimes = new List<TimeSpan>();
                for (int i = 0; i < iterations; i++)
                {
                    var startTime = DateTime.Now;
                    try
                    {
                        using var app = WordFactory.Open(tempPath);
                        // 简单操作
                        var _ = app.ActiveDocument.Paragraphs.Count;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Open测试出错: {ex.Message}");
                    }
                    var endTime = DateTime.Now;
                    openTimes.Add(endTime - startTime);
                }

                result.OpenAverageTime = openTimes.Average(t => t.TotalMilliseconds);
            }
            finally
            {
                // 清理测试文件
                if (File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"删除测试文件时出错: {ex.Message}");
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 生成API使用示例报告
        /// </summary>
        /// <returns>API使用示例报告</returns>
        public static string GenerateApiUsageReport()
        {
            var report = new StringBuilder();
            report.AppendLine("=== API使用示例报告 ===");
            report.AppendLine();

            report.AppendLine("1. WordFactory方法比较:");
            var comparisons = GetWordFactoryMethodComparisons();
            foreach (var comparison in comparisons)
            {
                report.AppendLine($"   方法: {comparison.MethodName}");
                report.AppendLine($"   用途: {comparison.Purpose}");
                report.AppendLine($"   特点: {comparison.Characteristics}");
                report.AppendLine($"   适用场景: {comparison.UseCase}");
                report.AppendLine();
            }

            report.AppendLine("2. 正确的API使用方式:");
            report.AppendLine("   // 正确：使用枚举常量");
            report.AppendLine("   selection.Font.Bold = 1;");
            report.AppendLine();
            report.AppendLine("   // 避免：使用特定语言的命令");
            report.AppendLine("   // app.CommandBars.Execute(\"Bold\"); // 可能在不同语言版本中失败");
            report.AppendLine();
            report.AppendLine("3. 资源管理最佳实践:");
            report.AppendLine("   // 使用using语句确保资源释放");
            report.AppendLine("   using var app = WordFactory.BlankWorkbook();");
            report.AppendLine("   // 使用app进行操作");
            report.AppendLine("   // 自动释放资源");
            report.AppendLine();
            report.AppendLine("=====================");

            return report.ToString();
        }
    }

    /// <summary>
    /// 性能比较结果类
    /// </summary>
    public class PerformanceComparisonResult
    {
        /// <summary>
        /// BlankWorkbook方法平均时间（毫秒）
        /// </summary>
        public double BlankWorkbookAverageTime { get; set; }

        /// <summary>
        /// Open方法平均时间（毫秒）
        /// </summary>
        public double OpenAverageTime { get; set; }

        /// <summary>
        /// CreateFrom方法平均时间（毫秒）
        /// </summary>
        public double CreateFromAverageTime { get; set; }
    }
}