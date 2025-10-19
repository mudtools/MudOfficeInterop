//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop;
using MudTools.OfficeInterop.Word;
using System.Text;

namespace DocumentAutomationProcessingSample
{
    /// <summary>
    /// 批量文档处理器类
    /// </summary>
    public class BatchDocumentProcessor
    {
        /// <summary>
        /// 批量处理目录中的文档
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="filePattern">文件匹配模式</param>
        /// <returns>处理结果</returns>
        public static BatchProcessingResult ProcessDocumentsInBatch(
            string inputDirectory,
            string outputDirectory,
            string filePattern = "*.docx")
        {
            var result = new BatchProcessingResult();

            try
            {
                // 确保输出目录存在
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // 获取所有匹配的文件
                var files = Directory.GetFiles(inputDirectory, filePattern);

                Console.WriteLine($"找到 {files.Length} 个文档需要处理");

                result.TotalFiles = files.Length;
                result.ProcessedFiles = new List<string>();
                result.FailedFiles = new List<string>();

                foreach (var file in files)
                {
                    try
                    {
                        Console.WriteLine($"正在处理: {Path.GetFileName(file)}");

                        // 处理单个文档
                        ProcessSingleDocument(file, outputDirectory);

                        result.ProcessedFiles.Add(file);
                        Console.WriteLine($"处理完成: {Path.GetFileName(file)}");
                    }
                    catch (Exception ex)
                    {
                        result.FailedFiles.Add(file);
                        Console.WriteLine($"处理 {Path.GetFileName(file)} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n批量处理完成:");
                Console.WriteLine($"成功处理: {result.ProcessedFiles.Count} 个文档");
                Console.WriteLine($"处理失败: {result.FailedFiles.Count} 个文档");
                Console.WriteLine($"总计处理: {files.Length} 个文档");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量处理过程中出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 处理单个文档
        /// </summary>
        /// <param name="inputFilePath">输入文件路径</param>
        /// <param name="outputDirectory">输出目录</param>
        private static void ProcessSingleDocument(string inputFilePath, string outputDirectory)
        {
            using var app = WordFactory.Open(inputFilePath);
            using var document = app.ActiveDocument;

            // 执行文档处理操作
            // 例如：标准化格式、更新字段、添加页眉页脚等
            StandardizeDocumentFormat(document);
            UpdateDocumentFields(document);
            AddHeaderFooter(document);

            // 生成输出文件路径
            var fileName = Path.GetFileNameWithoutExtension(inputFilePath);
            var outputFilePath = Path.Combine(outputDirectory, $"{fileName}_processed.docx");

            // 保存处理后的文档
            document.SaveAs(outputFilePath);
        }

        /// <summary>
        /// 标准化文档格式
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void StandardizeDocumentFormat(IWordDocument document)
        {
            try
            {
                // 标准化字体
                using var range = document.Range();
                range.Font.Name = "宋体";
                range.Font.Size = 12;

                // 标准化段落格式
                foreach (var paragraph in document.Paragraphs)
                {
                    using (paragraph)
                    {
                        paragraph.Format.LineSpacing = 1.5f; // 1.5倍行距
                        paragraph.Format.SpaceAfter = 12;    // 段后间距}
                        Console.WriteLine("  - 文档格式已标准化");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"标准化文档格式时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新文档字段
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void UpdateDocumentFields(IWordDocument document)
        {
            try
            {
                // 更新所有字段
                document.Range().Fields.Update();
                Console.WriteLine("  - 文档字段已更新");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新文档字段时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 添加页眉页脚
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private static void AddHeaderFooter(IWordDocument document)
        {
            try
            {
                // 添加页眉
                using var headerRange = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Text = "公司文档";
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                // 添加页脚（包含页码）
                using var footerRange = document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Fields.Add(footerRange, WdFieldType.wdFieldPage);
                footerRange.Text = " 第 页";
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                Console.WriteLine("  - 页眉页脚已添加");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加页眉页脚时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 批量处理特定类型的文档
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="documentType">文档类型</param>
        /// <returns>处理结果</returns>
        public static BatchProcessingResult ProcessDocumentsByType(
            string inputDirectory,
            string outputDirectory,
            DocumentType documentType)
        {
            string filePattern = documentType switch
            {
                DocumentType.Report => "*.docx",
                DocumentType.Contract => "*合同*.docx",
                DocumentType.Proposal => "*提案*.docx",
                DocumentType.Letter => "*信函*.docx",
                _ => "*.docx"
            };

            return ProcessDocumentsInBatch(inputDirectory, outputDirectory, filePattern);
        }

        /// <summary>
        /// 并行处理文档
        /// </summary>
        /// <param name="inputDirectory">输入目录</param>
        /// <param name="outputDirectory">输出目录</param>
        /// <param name="filePattern">文件匹配模式</param>
        /// <param name="maxDegreeOfParallelism">最大并行度</param>
        /// <returns>处理结果</returns>
        public static async Task<BatchProcessingResult> ProcessDocumentsInParallelAsync(
            string inputDirectory,
            string outputDirectory,
            string filePattern = "*.docx",
            int maxDegreeOfParallelism = 4)
        {
            var result = new BatchProcessingResult();

            try
            {
                // 确保输出目录存在
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // 获取所有匹配的文件
                var files = Directory.GetFiles(inputDirectory, filePattern);

                Console.WriteLine($"找到 {files.Length} 个文档需要处理");

                result.TotalFiles = files.Length;
                result.ProcessedFiles = new List<string>();
                result.FailedFiles = new List<string>();

                // 使用并行处理
                var parallelOptions = new ParallelOptions
                {
                    MaxDegreeOfParallelism = maxDegreeOfParallelism
                };

                var lockObject = new object();

                await Task.Run(() =>
                {
                    Parallel.ForEach(files, parallelOptions, file =>
                    {
                        try
                        {
                            Console.WriteLine($"正在处理: {Path.GetFileName(file)}");

                            // 处理单个文档
                            ProcessSingleDocument(file, outputDirectory);

                            lock (lockObject)
                            {
                                result.ProcessedFiles.Add(file);
                            }

                            Console.WriteLine($"处理完成: {Path.GetFileName(file)}");
                        }
                        catch (Exception ex)
                        {
                            lock (lockObject)
                            {
                                result.FailedFiles.Add(file);
                            }

                            Console.WriteLine($"处理 {Path.GetFileName(file)} 时出错: {ex.Message}");
                        }
                    });
                });

                Console.WriteLine($"\n并行批量处理完成:");
                Console.WriteLine($"成功处理: {result.ProcessedFiles.Count} 个文档");
                Console.WriteLine($"处理失败: {result.FailedFiles.Count} 个文档");
                Console.WriteLine($"总计处理: {files.Length} 个文档");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"并行批量处理过程中出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 生成处理报告
        /// </summary>
        /// <param name="result">处理结果</param>
        /// <returns>处理报告</returns>
        public static string GenerateProcessingReport(BatchProcessingResult result)
        {
            var report = new StringBuilder();
            report.AppendLine("=== 文档批量处理报告 ===");
            report.AppendLine($"处理时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"总文件数: {result.TotalFiles}");
            report.AppendLine($"成功处理: {result.ProcessedFiles.Count}");
            report.AppendLine($"处理失败: {result.FailedFiles.Count}");
            report.AppendLine($"成功率: {(result.TotalFiles > 0 ? (double)result.ProcessedFiles.Count / result.TotalFiles * 100 : 0):F2}%");

            if (result.FailedFiles.Any())
            {
                report.AppendLine("\n失败文件列表:");
                foreach (var file in result.FailedFiles)
                {
                    report.AppendLine($"  - {Path.GetFileName(file)}");
                }
            }

            if (!string.IsNullOrEmpty(result.ErrorMessage))
            {
                report.AppendLine($"\n错误信息: {result.ErrorMessage}");
            }

            report.AppendLine("========================");

            return report.ToString();
        }
    }

    /// <summary>
    /// 批量处理结果类
    /// </summary>
    public class BatchProcessingResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 总文件数
        /// </summary>
        public int TotalFiles { get; set; }

        /// <summary>
        /// 已处理文件列表
        /// </summary>
        public List<string> ProcessedFiles { get; set; } = new List<string>();

        /// <summary>
        /// 失败文件列表
        /// </summary>
        public List<string> FailedFiles { get; set; } = new List<string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 文档类型枚举
    /// </summary>
    public enum DocumentType
    {
        /// <summary>
        /// 报告
        /// </summary>
        Report,

        /// <summary>
        /// 合同
        /// </summary>
        Contract,

        /// <summary>
        /// 提案
        /// </summary>
        Proposal,

        /// <summary>
        /// 信函
        /// </summary>
        Letter,

        /// <summary>
        /// 其他
        /// </summary>
        Other
    }
}