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
    /// 文档处理器类
    /// </summary>
    public class DocumentProcessor
    {
        /// <summary>
        /// 创建示例文档
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="title">文档标题</param>
        /// <param name="content">文档内容</param>
        /// <returns>是否创建成功</returns>
        public static bool CreateSampleDocument(string filePath, string title, string content)
        {
            try
            {
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 设置文档标题
                document.Title = title;
                document.Author = "文档自动化系统";
                document.Subject = "示例文档";

                // 添加标题
                var titleRange = document.Range();
                titleRange.Text = $"{title}\n";
                titleRange.Font.Name = "微软雅黑";
                titleRange.Font.Size = 18;
                titleRange.Font.Bold = 1;
                titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleRange.ParagraphFormat.SpaceAfter = 24;

                // 添加内容
                var contentRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                contentRange.Text = content;
                contentRange.Font.Name = "宋体";
                contentRange.Font.Size = 12;

                // 保存文档
                document.SaveAs2(filePath);

                Console.WriteLine($"示例文档已创建: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建示例文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建批量示例文档
        /// </summary>
        /// <param name="directory">目录</param>
        /// <param name="documentInfos">文档信息列表</param>
        /// <returns>创建结果</returns>
        public static DocumentCreationResult CreateBatchSampleDocuments(
            string directory,
            List<DocumentInfo> documentInfos)
        {
            var result = new DocumentCreationResult();

            try
            {
                // 确保目录存在
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                Console.WriteLine($"开始创建 {documentInfos.Count} 个示例文档...");

                result.TotalDocuments = documentInfos.Count;
                result.CreatedDocuments = new List<string>();
                result.FailedDocuments = new List<string>();

                foreach (var docInfo in documentInfos)
                {
                    try
                    {
                        string filePath = Path.Combine(directory, docInfo.FileName);
                        bool success = CreateSampleDocument(filePath, docInfo.Title, docInfo.Content);

                        if (success)
                        {
                            result.CreatedDocuments.Add(filePath);
                            Console.WriteLine($"已创建: {docInfo.FileName}");
                        }
                        else
                        {
                            result.FailedDocuments.Add(filePath);
                            Console.WriteLine($"创建失败: {docInfo.FileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        string filePath = Path.Combine(directory, docInfo.FileName);
                        result.FailedDocuments.Add(filePath);
                        Console.WriteLine($"创建 {docInfo.FileName} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n批量创建完成:");
                Console.WriteLine($"成功创建: {result.CreatedDocuments.Count} 个文档");
                Console.WriteLine($"创建失败: {result.FailedDocuments.Count} 个文档");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量创建示例文档时出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 处理文档内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="operations">操作列表</param>
        /// <returns>处理结果</returns>
        public static ContentProcessingResult ProcessDocumentContent(
            IWordDocument document,
            List<ContentOperation> operations)
        {
            var result = new ContentProcessingResult();

            try
            {
                Console.WriteLine("开始处理文档内容...");

                result.TotalOperations = operations.Count;
                result.CompletedOperations = new List<ContentOperation>();
                result.FailedOperations = new List<ContentOperation>();

                foreach (var operation in operations)
                {
                    try
                    {
                        switch (operation.Type)
                        {
                            case ContentOperationType.ReplaceText:
                                ReplaceText(document, operation.Parameters);
                                break;
                            case ContentOperationType.InsertText:
                                InsertText(document, operation.Parameters);
                                break;
                            case ContentOperationType.DeleteText:
                                DeleteText(document, operation.Parameters);
                                break;
                            case ContentOperationType.FormatText:
                                FormatText(document, operation.Parameters);
                                break;
                        }

                        result.CompletedOperations.Add(operation);
                        Console.WriteLine($"已完成操作: {operation.Description}");
                    }
                    catch (Exception ex)
                    {
                        result.FailedOperations.Add(operation);
                        Console.WriteLine($"执行操作 {operation.Description} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"\n内容处理完成:");
                Console.WriteLine($"成功执行: {result.CompletedOperations.Count} 个操作");
                Console.WriteLine($"执行失败: {result.FailedOperations.Count} 个操作");

                result.Success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理文档内容时出错: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// 替换文本
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="parameters">参数字典</param>
        private static void ReplaceText(IWordDocument document, Dictionary<string, string> parameters)
        {
            if (parameters.ContainsKey("FindText") && parameters.ContainsKey("ReplaceText"))
            {
                var find = document.Range().Find;
                find.ClearFormatting();
                find.Text = parameters["FindText"];
                find.Replacement.Text = parameters["ReplaceText"];
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.Execute(replace: WdReplace.wdReplaceAll);
            }
        }

        /// <summary>
        /// 插入文本
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="parameters">参数字典</param>
        private static void InsertText(IWordDocument document, Dictionary<string, string> parameters)
        {
            if (parameters.ContainsKey("Position") && parameters.ContainsKey("Text"))
            {
                int position = int.Parse(parameters["Position"]);
                string text = parameters["Text"];

                var range = document.Range(position, position);
                range.Text = text;
            }
        }

        /// <summary>
        /// 删除文本
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="parameters">参数字典</param>
        private static void DeleteText(IWordDocument document, Dictionary<string, string> parameters)
        {
            if (parameters.ContainsKey("Start") && parameters.ContainsKey("End"))
            {
                int start = int.Parse(parameters["Start"]);
                int end = int.Parse(parameters["End"]);

                var range = document.Range(start, end);
                range.Text = "";
            }
        }

        /// <summary>
        /// 格式化文本
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="parameters">参数字典</param>
        private static void FormatText(IWordDocument document, Dictionary<string, string> parameters)
        {
            if (parameters.ContainsKey("Start") && parameters.ContainsKey("End"))
            {
                int start = int.Parse(parameters["Start"]);
                int end = int.Parse(parameters["End"]);

                var range = document.Range(start, end);

                // 应用格式设置
                if (parameters.ContainsKey("FontName"))
                {
                    range.Font.Name = parameters["FontName"];
                }

                if (parameters.ContainsKey("FontSize"))
                {
                    range.Font.Size = float.Parse(parameters["FontSize"]);
                }

                if (parameters.ContainsKey("Bold") && parameters["Bold"] == "true")
                {
                    range.Font.Bold = true;
                }

                if (parameters.ContainsKey("Italic") && parameters["Italic"] == "true")
                {
                    range.Font.Italic = true;
                }
            }
        }

        /// <summary>
        /// 分析文档统计信息
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>文档统计信息</returns>
        public static DocumentStatistics AnalyzeDocumentStatistics(IWordDocument document)
        {
            var statistics = new DocumentStatistics();

            try
            {
                // 获取基本统计信息
                statistics.PageCount = document.Range().Paragraphs.Count > 0 ? document.ComputeStatistics(WdStatistic.wdStatisticPages) : 0;
                statistics.WordCount = document.Range().Paragraphs.Count > 0 ? document.ComputeStatistics(WdStatistic.wdStatisticWords) : 0;
                statistics.CharacterCount = document.Range().Paragraphs.Count > 0 ? document.ComputeStatistics(WdStatistic.wdStatisticCharacters) : 0;
                statistics.ParagraphCount = document.Paragraphs.Count;
                statistics.TableCount = document.Tables.Count;
                statistics.ImageFieldCount = document.Shapes.Count + document.InlineShapes.Count;

                // 分析段落格式
                foreach (var paragraph in document.Paragraphs)
                {
                    var format = paragraph.Format;
                    if (format.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        statistics.CenterAlignedParagraphs++;
                    }
                    else if (format.Alignment == WdParagraphAlignment.wdAlignParagraphLeft)
                    {
                        statistics.LeftAlignedParagraphs++;
                    }
                    else if (format.Alignment == WdParagraphAlignment.wdAlignParagraphRight)
                    {
                        statistics.RightAlignedParagraphs++;
                    }
                }

                Console.WriteLine("文档统计信息分析完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"分析文档统计信息时出错: {ex.Message}");
                statistics.ErrorMessage = ex.Message;
            }

            return statistics;
        }

        /// <summary>
        /// 生成文档报告
        /// </summary>
        /// <param name="statistics">文档统计信息</param>
        /// <returns>文档报告</returns>
        public static string GenerateDocumentReport(DocumentStatistics statistics)
        {
            var report = new StringBuilder();
            report.AppendLine("=== 文档分析报告 ===");
            report.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine();
            report.AppendLine("基本统计:");
            report.AppendLine($"  页数: {statistics.PageCount}");
            report.AppendLine($"  字数: {statistics.WordCount}");
            report.AppendLine($"  字符数: {statistics.CharacterCount}");
            report.AppendLine($"  段落数: {statistics.ParagraphCount}");
            report.AppendLine($"  表格数: {statistics.TableCount}");
            report.AppendLine($"  图片和形状数: {statistics.ImageFieldCount}");
            report.AppendLine();
            report.AppendLine("段落对齐:");
            report.AppendLine($"  左对齐段落: {statistics.LeftAlignedParagraphs}");
            report.AppendLine($"  居中段落: {statistics.CenterAlignedParagraphs}");
            report.AppendLine($"  右对齐段落: {statistics.RightAlignedParagraphs}");

            if (!string.IsNullOrEmpty(statistics.ErrorMessage))
            {
                report.AppendLine($"\n错误信息: {statistics.ErrorMessage}");
            }

            report.AppendLine("==================");

            return report.ToString();
        }

        /// <summary>
        /// 优化文档性能
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>是否优化成功</returns>
        public static bool OptimizeDocumentPerformance(IWordDocument document)
        {
            try
            {
                // 禁用屏幕更新以提高性能
                var application = document.Application;
                bool oldScreenUpdating = application.ScreenUpdating;
                application.ScreenUpdating = false;

                // 禁用事件处理
                bool oldEvents = application.Events;
                application.Events = false;

                // 执行优化操作
                // 例如：压缩图片、删除未使用的格式等

                // 恢复设置
                application.ScreenUpdating = oldScreenUpdating;
                application.Events = oldEvents;

                Console.WriteLine("文档性能优化完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"优化文档性能时出错: {ex.Message}");
                return false;
            }
        }
    }

    /// <summary>
    /// 文档信息类
    /// </summary>
    public class DocumentInfo
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 内容
        /// </summary>
        public string Content { get; set; }
    }

    /// <summary>
    /// 文档创建结果类
    /// </summary>
    public class DocumentCreationResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 总文档数
        /// </summary>
        public int TotalDocuments { get; set; }

        /// <summary>
        /// 已创建文档列表
        /// </summary>
        public List<string> CreatedDocuments { get; set; } = new List<string>();

        /// <summary>
        /// 失败文档列表
        /// </summary>
        public List<string> FailedDocuments { get; set; } = new List<string>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 内容处理结果类
    /// </summary>
    public class ContentProcessingResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 总操作数
        /// </summary>
        public int TotalOperations { get; set; }

        /// <summary>
        /// 已完成操作列表
        /// </summary>
        public List<ContentOperation> CompletedOperations { get; set; } = new List<ContentOperation>();

        /// <summary>
        /// 失败操作列表
        /// </summary>
        public List<ContentOperation> FailedOperations { get; set; } = new List<ContentOperation>();

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 内容操作类
    /// </summary>
    public class ContentOperation
    {
        /// <summary>
        /// 操作类型
        /// </summary>
        public ContentOperationType Type { get; set; }

        /// <summary>
        /// 参数字典
        /// </summary>
        public Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 操作描述
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// 内容操作类型枚举
    /// </summary>
    public enum ContentOperationType
    {
        /// <summary>
        /// 替换文本
        /// </summary>
        ReplaceText,

        /// <summary>
        /// 插入文本
        /// </summary>
        InsertText,

        /// <summary>
        /// 删除文本
        /// </summary>
        DeleteText,

        /// <summary>
        /// 格式化文本
        /// </summary>
        FormatText
    }

    /// <summary>
    /// 文档统计信息类
    /// </summary>
    public class DocumentStatistics
    {
        /// <summary>
        /// 页数
        /// </summary>
        public int PageCount { get; set; }

        /// <summary>
        /// 字数
        /// </summary>
        public int WordCount { get; set; }

        /// <summary>
        /// 字符数
        /// </summary>
        public int CharacterCount { get; set; }

        /// <summary>
        /// 段落数
        /// </summary>
        public int ParagraphCount { get; set; }

        /// <summary>
        /// 表格数
        /// </summary>
        public int TableCount { get; set; }

        /// <summary>
        /// 图片和形状数
        /// </summary>
        public int ImageFieldCount { get; set; }

        /// <summary>
        /// 左对齐段落数
        /// </summary>
        public int LeftAlignedParagraphs { get; set; }

        /// <summary>
        /// 居中段落数
        /// </summary>
        public int CenterAlignedParagraphs { get; set; }

        /// <summary>
        /// 右对齐段落数
        /// </summary>
        public int RightAlignedParagraphs { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}