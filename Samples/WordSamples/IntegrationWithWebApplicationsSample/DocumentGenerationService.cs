using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrationWithWebApplicationsSample
{
    /// <summary>
    /// 文档生成服务类
    /// </summary>
    public class DocumentGenerationService
    {
        /// <summary>
        /// 日志记录器
        /// </summary>
        private readonly ILogger _logger;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        public DocumentGenerationService(ILogger logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 生成文档
        /// </summary>
        /// <param name="request">文档请求</param>
        /// <returns>生成结果</returns>
        public async Task<DocumentGenerationResult> GenerateDocumentAsync(DocumentRequest request)
        {
            try
            {
                _logger?.LogInformation("开始生成文档: {Title}", request.Title);

                // 创建文档内容
                var contentBuilder = new StringBuilder();
                contentBuilder.AppendLine(request.Title);
                contentBuilder.AppendLine($"作者: {request.Author}");
                contentBuilder.AppendLine(new string('=', 50));
                contentBuilder.AppendLine();

                if (!string.IsNullOrEmpty(request.Content))
                {
                    contentBuilder.AppendLine(request.Content);
                    contentBuilder.AppendLine();
                }

                if (request.Sections != null && request.Sections.Count > 0)
                {
                    foreach (var section in request.Sections)
                    {
                        contentBuilder.AppendLine(section.Heading);
                        contentBuilder.AppendLine(new string('-', section.Heading.Length));
                        contentBuilder.AppendLine(section.Text);
                        contentBuilder.AppendLine();
                    }
                }

                // 使用WordFactory创建文档
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加内容到文档
                document.Range().Text = contentBuilder.ToString();

                // 格式化标题
                if (!string.IsNullOrEmpty(request.Title))
                {
                    var titleRange = document.Range(0, request.Title.Length);
                    titleRange.Font.Size = 18;
                    titleRange.Font.Bold = 1;
                    titleRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // 保存文档
                var fileName = $"Document_{Guid.NewGuid()}.docx";
                var filePath = Path.Combine(Path.GetTempPath(), fileName);
                document.SaveAs2(filePath);

                _logger?.LogInformation("文档生成完成: {Title}", request.Title);

                return new DocumentGenerationResult
                {
                    FilePath = filePath,
                    Success = true,
                    Message = "文档生成成功"
                };
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "生成文档时出错: {Message}", ex.Message);

                return new DocumentGenerationResult
                {
                    Success = false,
                    Message = $"生成文档时发生错误: {ex.Message}",
                    ErrorMessage = ex.ToString()
                };
            }
        }

        /// <summary>
        /// 从模板生成文档
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="request">文档请求</param>
        /// <returns>生成结果</returns>
        public async Task<DocumentGenerationResult> GenerateDocumentFromTemplateAsync(string templatePath, DocumentRequest request)
        {
            try
            {
                _logger?.LogInformation("开始从模板生成文档: {Title}", request.Title);

                // 使用WordFactory从模板创建文档
                using var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;

                // 替换模板中的占位符
                ReplacePlaceholders(document, request);

                // 保存文档
                var fileName = $"Document_{Guid.NewGuid()}.docx";
                var filePath = Path.Combine(Path.GetTempPath(), fileName);
                document.SaveAs2(filePath);

                _logger?.LogInformation("从模板生成文档完成: {Title}", request.Title);

                return new DocumentGenerationResult
                {
                    FilePath = filePath,
                    Success = true,
                    Message = "文档生成成功"
                };
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "从模板生成文档时出错: {Message}", ex.Message);

                return new DocumentGenerationResult
                {
                    Success = false,
                    Message = $"从模板生成文档时发生错误: {ex.Message}",
                    ErrorMessage = ex.ToString()
                };
            }
        }

        /// <summary>
        /// 替换模板中的占位符
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="request">文档请求</param>
        private void ReplacePlaceholders(IWordDocument document, DocumentRequest request)
        {
            // 替换标题占位符
            if (!string.IsNullOrEmpty(request.Title))
            {
                ReplaceText(document, "{{Title}}", request.Title);
            }

            // 替换作者占位符
            if (!string.IsNullOrEmpty(request.Author))
            {
                ReplaceText(document, "{{Author}}", request.Author);
            }

            // 替换内容占位符
            if (!string.IsNullOrEmpty(request.Content))
            {
                ReplaceText(document, "{{Content}}", request.Content);
            }

            // 替换日期占位符
            ReplaceText(document, "{{Date}}", DateTime.Now.ToString("yyyy-MM-dd"));

            // 替换自定义字段占位符
            if (request.CustomFields != null)
            {
                foreach (var field in request.CustomFields)
                {
                    ReplaceText(document, $"{{{{{field.Key}}}}}", field.Value);
                }
            }
        }

        /// <summary>
        /// 替换文档中的文本
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="findText">查找文本</param>
        /// <param name="replaceText">替换文本</param>
        private void ReplaceText(IWordDocument document, string findText, string replaceText)
        {
            var range = document.Range();
            var find = range.Find;
            find.ClearFormatting();
            find.Text = findText;
            find.Replacement.Text = replaceText;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindContinue;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;
            find.Execute(Replace: WdReplace.wdReplaceAll);
        }

        /// <summary>
        /// 批量生成文档
        /// </summary>
        /// <param name="requests">文档请求列表</param>
        /// <returns>生成结果列表</returns>
        public async Task<List<DocumentGenerationResult>> GenerateDocumentsAsync(List<DocumentRequest> requests)
        {
            var results = new List<DocumentGenerationResult>();

            foreach (var request in requests)
            {
                try
                {
                    var result = await GenerateDocumentAsync(request);
                    results.Add(result);
                }
                catch (Exception ex)
                {
                    results.Add(new DocumentGenerationResult
                    {
                        Success = false,
                        Message = $"处理文档 '{request.Title}' 时发生错误",
                        ErrorMessage = ex.ToString()
                    });
                }
            }

            return results;
        }

        /// <summary>
        /// 转换文档格式
        /// </summary>
        /// <param name="inputPath">输入文档路径</param>
        /// <param name="outputFormat">输出格式</param>
        /// <returns>转换结果</returns>
        public async Task<DocumentGenerationResult> ConvertDocumentAsync(string inputPath, WdSaveFormat outputFormat)
        {
            try
            {
                _logger?.LogInformation("开始转换文档格式: {InputPath}", inputPath);

                // 打开文档
                using var app = WordFactory.Open(inputPath);
                var document = app.ActiveDocument;

                // 生成输出文件路径
                var fileName = Path.GetFileNameWithoutExtension(inputPath);
                string extension = GetExtensionForFormat(outputFormat);
                var outputFilePath = Path.Combine(Path.GetTempPath(), $"{fileName}{extension}");

                // 保存为指定格式
                document.SaveAs2(outputFilePath, outputFormat);

                _logger?.LogInformation("文档格式转换完成: {OutputPath}", outputFilePath);

                return new DocumentGenerationResult
                {
                    FilePath = outputFilePath,
                    Success = true,
                    Message = "文档格式转换成功"
                };
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "转换文档格式时出错: {Message}", ex.Message);

                return new DocumentGenerationResult
                {
                    Success = false,
                    Message = $"转换文档格式时发生错误: {ex.Message}",
                    ErrorMessage = ex.ToString()
                };
            }
        }

        /// <summary>
        /// 根据格式获取文件扩展名
        /// </summary>
        /// <param name="format">保存格式</param>
        /// <returns>文件扩展名</returns>
        private string GetExtensionForFormat(WdSaveFormat format)
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
    }

    /// <summary>
    /// 文档请求类
    /// </summary>
    public class DocumentRequest
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 内容
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// 文档章节列表
        /// </summary>
        public List<DocumentSection> Sections { get; set; } = new List<DocumentSection>();

        /// <summary>
        /// 自定义字段
        /// </summary>
        public Dictionary<string, string> CustomFields { get; set; } = new Dictionary<string, string>();
    }

    /// <summary>
    /// 文档章节类
    /// </summary>
    public class DocumentSection
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Heading { get; set; }

        /// <summary>
        /// 文本
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// 是否重要
        /// </summary>
        public bool IsImportant { get; set; }
    }

    /// <summary>
    /// 文档生成结果类
    /// </summary>
    public class DocumentGenerationResult
    {
        /// <summary>
        /// 文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 日志记录器接口
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// 记录信息日志
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        void LogInformation(string message, params object[] args);

        /// <summary>
        /// 记录错误日志
        /// </summary>
        /// <param name="exception">异常</param>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        void LogError(Exception exception, string message, params object[] args);
    }

    /// <summary>
    /// 控制台日志记录器实现
    /// </summary>
    public class ConsoleLogger : ILogger
    {
        /// <summary>
        /// 记录信息日志
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void LogInformation(string message, params object[] args)
        {
            Console.WriteLine($"[INFO] {string.Format(message, args)}");
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        /// <param name="exception">异常</param>
        /// <param name="message">消息</param>
        /// <param name="args">参数</param>
        public void LogError(Exception exception, string message, params object[] args)
        {
            Console.WriteLine($"[ERROR] {string.Format(message, args)}: {exception.Message}");
        }
    }
}