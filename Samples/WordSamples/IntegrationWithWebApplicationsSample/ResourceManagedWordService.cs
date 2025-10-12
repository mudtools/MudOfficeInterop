using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace IntegrationWithWebApplicationsSample
{
    /// <summary>
    /// 资源管理的Word服务类
    /// </summary>
    /// <remarks>
    /// 在Web环境中，正确的资源管理对系统稳定性至关重要
    /// </remarks>
    public class ResourceManagedWordService : IDisposable
    {
        private readonly object _lockObject = new object();
        private IWordApplication _wordApp;
        private bool _disposed = false;

        /// <summary>
        /// 构造函数
        /// </summary>
        public ResourceManagedWordService()
        {
            // 初始化时创建Word应用实例
            InitializeWordApplication();
        }

        /// <summary>
        /// 初始化Word应用程序实例
        /// </summary>
        private void InitializeWordApplication()
        {
            lock (_lockObject)
            {
                if (_wordApp == null)
                {
                    try
                    {
                        _wordApp = WordFactory.BlankWorkbook();
                        _wordApp.Visible = false; // Web环境中隐藏界面
                        _wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone; // 禁用警告对话框
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("无法初始化Word应用程序", ex);
                    }
                }
            }
        }

        /// <summary>
        /// 生成文档
        /// </summary>
        /// <param name="content">文档内容</param>
        /// <returns>文档路径</returns>
        public string GenerateDocument(string content)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ResourceManagedWordService));

            lock (_lockObject)
            {
                try
                {
                    // 创建新文档
                    var document = _wordApp.Documents.Add();

                    try
                    {
                        // 处理文档
                        document.Range().Text = content;

                        // 保存到临时文件
                        var tempPath = Path.GetTempFileName().Replace(".tmp", ".docx");
                        document.SaveAs2(tempPath);

                        return tempPath;
                    }
                    finally
                    {
                        // 关闭文档但不退出Word应用
                        document.Close(WdSaveOptions.wdDoNotSaveChanges);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("生成文档时出错", ex);
                }
            }
        }

        /// <summary>
        /// 从模板生成文档
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="data">文档数据</param>
        /// <returns>文档路径</returns>
        public string GenerateDocumentFromTemplate(string templatePath, object data)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ResourceManagedWordService));

            lock (_lockObject)
            {
                try
                {
                    // 从模板创建文档
                    var document = _wordApp.Documents.Add(templatePath);

                    try
                    {
                        // 处理文档（填充数据、格式化等）
                        FillTemplateData(document, data);
                        ApplyFormatting(document);

                        // 保存到临时文件
                        var tempPath = Path.GetTempFileName().Replace(".tmp", ".docx");
                        document.SaveAs2(tempPath);

                        return tempPath;
                    }
                    finally
                    {
                        // 关闭文档但不退出Word应用
                        document.Close(WdSaveOptions.wdDoNotSaveChanges);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("从模板生成文档时出错", ex);
                }
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
        /// 执行文档操作
        /// </summary>
        /// <param name="operation">文档操作</param>
        public void ExecuteDocumentOperation(Action<IWordDocument> operation)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ResourceManagedWordService));

            lock (_lockObject)
            {
                try
                {
                    // 创建新文档
                    var document = _wordApp.Documents.Add();

                    try
                    {
                        // 执行操作
                        operation(document);

                        // 保存到临时文件
                        var tempPath = Path.GetTempFileName().Replace(".tmp", ".docx");
                        document.SaveAs2(tempPath);
                    }
                    finally
                    {
                        // 关闭文档但不退出Word应用
                        document.Close(WdSaveOptions.wdDoNotSaveChanges);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("执行文档操作时出错", ex);
                }
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        /// <param name="disposing">是否正在 disposing</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    lock (_lockObject)
                    {
                        try
                        {
                            // 关闭所有文档
                            for (int i = _wordApp.Documents.Count; i > 0; i--)
                            {
                                var document = _wordApp.Documents[i];
                                document.Close(WdSaveOptions.wdDoNotSaveChanges);
                            }

                            // 退出Word应用
                            _wordApp.Quit();
                        }
                        catch (Exception ex)
                        {
                            // 记录日志但不抛出异常
                            Console.WriteLine($"关闭Word应用时出错: {ex.Message}");
                        }
                        finally
                        {
                            _wordApp = null;
                        }
                    }
                }

                _disposed = true;
            }
        }

        /// <summary>
        /// 析构函数
        /// </summary>
        ~ResourceManagedWordService()
        {
            Dispose(false);
        }
    }
}