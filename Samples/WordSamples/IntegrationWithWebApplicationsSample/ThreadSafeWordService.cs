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
    /// 线程安全的Word服务类
    /// </summary>
    /// <remarks>
    /// Office应用程序是单线程的，需要确保在STA线程模型中使用
    /// </remarks>
    public class ThreadSafeWordService
    {
        /// <summary>
        /// 处理文档
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">数据对象</param>
        /// <returns>文档路径</returns>
        public async Task<string> ProcessDocumentAsync(string templatePath, object data)
        {
            // 在STA线程中执行
            var task = Task.Factory.StartNew(() =>
            {
                // 设置线程为STA模式
                Thread.CurrentThread.SetApartmentState(ApartmentState.STA);

                return ProcessDocumentInternal(templatePath, data);
            }, TaskCreationOptions.LongRunning);

            return await task;
        }

        /// <summary>
        /// 处理文档内部逻辑
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">数据对象</param>
        /// <returns>文档路径</returns>
        private string ProcessDocumentInternal(string templatePath, object data)
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
        /// 在STA线程中执行操作
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <returns>任务</returns>
        public async Task ExecuteInStaThreadAsync(Action action)
        {
            var task = Task.Factory.StartNew(() =>
            {
                Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                action();
            }, TaskCreationOptions.LongRunning);

            await task;
        }

        /// <summary>
        /// 在STA线程中执行操作并返回结果
        /// </summary>
        /// <typeparam name="T">返回类型</typeparam>
        /// <param name="func">要执行的函数</param>
        /// <returns>任务结果</returns>
        public async Task<T> ExecuteInStaThreadAsync<T>(Func<T> func)
        {
            var task = Task.Factory.StartNew(() =>
            {
                Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                return func();
            }, TaskCreationOptions.LongRunning);

            return await task;
        }

        /// <summary>
        /// 批量处理文档
        /// </summary>
        /// <param name="documents">文档处理请求列表</param>
        /// <returns>处理结果</returns>
        public async Task<List<DocumentProcessingResult>> ProcessDocumentsAsync(List<DocumentProcessingRequest> documents)
        {
            var results = new List<DocumentProcessingResult>();

            foreach (var request in documents)
            {
                try
                {
                    var result = await ProcessDocumentAsync(request.TemplatePath, request.Data);
                    results.Add(new DocumentProcessingResult
                    {
                        DocumentName = request.DocumentName,
                        OutputPath = result,
                        Success = true
                    });
                }
                catch (Exception ex)
                {
                    results.Add(new DocumentProcessingResult
                    {
                        DocumentName = request.DocumentName,
                        ErrorMessage = ex.Message,
                        Success = false
                    });
                }
            }

            return results;
        }
    }

    /// <summary>
    /// 文档处理请求类
    /// </summary>
    public class DocumentProcessingRequest
    {
        /// <summary>
        /// 文档名称
        /// </summary>
        public string DocumentName { get; set; }

        /// <summary>
        /// 模板路径
        /// </summary>
        public string TemplatePath { get; set; }

        /// <summary>
        /// 数据对象
        /// </summary>
        public object Data { get; set; }
    }

    /// <summary>
    /// 文档处理结果类
    /// </summary>
    public class DocumentProcessingResult
    {
        /// <summary>
        /// 文档名称
        /// </summary>
        public string DocumentName { get; set; }

        /// <summary>
        /// 输出路径
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}