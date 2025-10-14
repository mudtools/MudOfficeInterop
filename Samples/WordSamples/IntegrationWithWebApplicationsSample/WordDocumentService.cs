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
    /// <summary>
    /// Word文档服务类
    /// </summary>
    /// <remarks>
    /// 在Web应用中使用Office COM组件需要特别注意线程模型和安全性问题
    /// </remarks>
    public class WordDocumentService
    {
        /// <summary>
        /// 使用信号量控制并发访问
        /// </summary>
        private static readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

        /// <summary>
        /// 创建文档
        /// </summary>
        /// <param name="content">文档内容</param>
        /// <returns>文档路径</returns>
        public async Task<string> CreateDocumentAsync(string content)
        {
            await _semaphore.WaitAsync();
            try
            {
                // 设置线程为STA模式（如果在新线程中运行）
                using var app = WordFactory.BlankWorkbook();
                var document = app.ActiveDocument;

                // 添加内容
                document.Range().Text = content;

                // 保存文档
                var fileName = $"Document_{Guid.NewGuid()}.docx";
                var filePath = Path.Combine(Path.GetTempPath(), fileName);
                document.SaveAs(filePath);

                return filePath;
            }
            finally
            {
                _semaphore.Release();
            }
        }

        /// <summary>
        /// 更安全的实现方式 - 使用独立进程
        /// </summary>
        /// <param name="content">文档内容</param>
        /// <returns>文档路径</returns>
        public async Task<string> CreateDocumentInProcessAsync(string content)
        {
            // 创建独立进程来处理Word文档
            // 注意：在实际应用中，这需要一个独立的文档处理程序
            var processInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "DocumentProcessor.exe",
                Arguments = $"\"{content}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            // 在实际应用中，您需要创建一个独立的文档处理程序
            // 这里我们只是演示概念
            Console.WriteLine("模拟独立进程处理Word文档...");
            await Task.Delay(1000); // 模拟处理时间

            var fileName = $"Document_{Guid.NewGuid()}.docx";
            var filePath = Path.Combine(Path.GetTempPath(), fileName);
            return filePath;
        }

        /// <summary>
        /// 从模板创建文档
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">文档数据</param>
        /// <returns>文档路径</returns>
        public async Task<string> CreateDocumentFromTemplateAsync(string templatePath, object data)
        {
            await _semaphore.WaitAsync();
            try
            {
                using var app = WordFactory.CreateFrom(templatePath);
                var document = app.ActiveDocument;

                // 处理文档（填充数据、格式化等）
                FillTemplateData(document, data);
                ApplyFormatting(document);

                // 保存文档
                var outputPath = Path.GetTempFileName().Replace(".tmp", ".docx");
                document.SaveAs(outputPath);

                return outputPath;
            }
            finally
            {
                _semaphore.Release();
            }
        }

        /// <summary>
        /// 填充模板数据
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="data">数据对象</param>
        private void FillTemplateData(IWordDocument document, object data)
        {
            // 实现模板数据填充逻辑
            // 这里简化处理
            var range = document.Range();
            var text = range.Text;

            // 替换占位符（示例）
            if (data is IDictionary<string, string> keyValuePairs)
            {
                foreach (var pair in keyValuePairs)
                {
                    text = text.Replace($"{{{pair.Key}}}", pair.Value);
                }
            }

            range.Text = text;
        }

        /// <summary>
        /// 应用标准格式化
        /// </summary>
        /// <param name="document">Word文档对象</param>
        private void ApplyFormatting(IWordDocument document)
        {
            // 应用标准格式化
            var range = document.Range();
            range.Font.Name = "宋体";
            range.Font.Size = 12;
        }

        /// <summary>
        /// 转换文档格式
        /// </summary>
        /// <param name="inputPath">输入文档路径</param>
        /// <param name="outputFormat">输出格式</param>
        /// <returns>转换后的文档路径</returns>
        public async Task<string> ConvertDocumentAsync(string inputPath, WdSaveFormat outputFormat)
        {
            await _semaphore.WaitAsync();
            try
            {
                using var app = WordFactory.Open(inputPath);
                var document = app.ActiveDocument;

                // 生成输出文件路径
                var fileName = Path.GetFileNameWithoutExtension(inputPath);
                string extension = GetExtensionForFormat(outputFormat);
                var outputFilePath = Path.Combine(Path.GetTempPath(), $"{fileName}{extension}");

                // 保存为指定格式
                document.SaveAs(outputFilePath, outputFormat);

                return outputFilePath;
            }
            finally
            {
                _semaphore.Release();
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
                WdSaveFormat.wdFormatFlatXML => ".xml",
                WdSaveFormat.wdFormatPDF => ".pdf",
                WdSaveFormat.wdFormatRTF => ".rtf",
                WdSaveFormat.wdFormatFilteredHTML => ".htm",
                _ => ".docx"
            };
        }

        /// <summary>
        /// 批量处理文档
        /// </summary>
        /// <param name="documents">文档列表</param>
        /// <returns>处理结果</returns>
        public async Task<BatchProcessingResult> ProcessDocumentsAsync(List<DocumentInfo> documents)
        {
            var result = new BatchProcessingResult
            {
                TotalDocuments = documents.Count,
                ProcessedDocuments = new List<string>(),
                FailedDocuments = new List<string>()
            };

            foreach (var doc in documents)
            {
                try
                {
                    string outputPath = await CreateDocumentAsync(doc.Content);
                    result.ProcessedDocuments.Add(outputPath);
                    Console.WriteLine($"文档处理完成: {doc.Name}");
                }
                catch (Exception ex)
                {
                    result.FailedDocuments.Add(doc.Name);
                    Console.WriteLine($"文档处理失败 {doc.Name}: {ex.Message}");
                }
            }

            return result;
        }
    }

    /// <summary>
    /// 文档信息类
    /// </summary>
    public class DocumentInfo
    {
        /// <summary>
        /// 文档名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 文档内容
        /// </summary>
        public string Content { get; set; }
    }

    /// <summary>
    /// 批量处理结果类
    /// </summary>
    public class BatchProcessingResult
    {
        /// <summary>
        /// 总文档数
        /// </summary>
        public int TotalDocuments { get; set; }

        /// <summary>
        /// 已处理文档列表
        /// </summary>
        public List<string> ProcessedDocuments { get; set; } = new List<string>();

        /// <summary>
        /// 失败文档列表
        /// </summary>
        public List<string> FailedDocuments { get; set; } = new List<string>();
    }
}