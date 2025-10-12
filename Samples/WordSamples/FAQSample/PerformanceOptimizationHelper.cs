using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FAQSample
{
    /// <summary>
    /// 性能优化帮助类
    /// </summary>
    public class PerformanceOptimizationHelper
    {
        /// <summary>
        /// 高性能文档处理器
        /// </summary>
        public class HighPerformanceDocumentProcessor : IDisposable
        {
            private readonly IWordApplication _application;
            private bool _disposed = false;

            /// <summary>
            /// 构造函数
            /// </summary>
            public HighPerformanceDocumentProcessor()
            {
                _application = WordFactory.BlankWorkbook();
                OptimizeForPerformance();
            }

            /// <summary>
            /// 为性能优化Word应用程序
            /// </summary>
            private void OptimizeForPerformance()
            {
                _application.Visible = false; // 隐藏界面
                _application.ScreenUpdating = false; // 禁用屏幕更新
                _application.DisplayAlerts = WdAlertLevel.wdAlertsNone; // 禁用警告
            }

            /// <summary>
            /// 批量处理文档
            /// </summary>
            /// <param name="documentPaths">文档路径列表</param>
            /// <param name="processAction">处理操作</param>
            /// <returns>处理结果</returns>
            public BatchProcessingResult ProcessDocuments(
                List<string> documentPaths,
                Action<IWordDocument> processAction)
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(HighPerformanceDocumentProcessor));

                var result = new BatchProcessingResult
                {
                    TotalDocuments = documentPaths.Count,
                    ProcessedDocuments = new List<string>(),
                    FailedDocuments = new List<string>()
                };

                foreach (var path in documentPaths)
                {
                    try
                    {
                        var document = _application.Documents.Open(path);
                        try
                        {
                            processAction(document);
                            result.ProcessedDocuments.Add(path);
                        }
                        finally
                        {
                            document.Close(WdSaveOptions.wdDoNotSaveChanges);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedDocuments.Add(path);
                        Console.WriteLine($"处理文档 {path} 时出错: {ex.Message}");
                    }
                }

                return result;
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
            /// <param name="disposing">是否正在disposing</param>
            protected virtual void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        try
                        {
                            // 恢复设置
                            _application.ScreenUpdating = true;
                            _application.DisplayAlerts = WdAlertLevel.wdAlertsAll;

                            // 退出应用程序
                            _application.Quit();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"释放Word应用程序时出错: {ex.Message}");
                        }
                    }

                    _disposed = true;
                }
            }
        }

        /// <summary>
        /// 创建优化的Word应用程序
        /// </summary>
        /// <returns>优化的Word应用程序</returns>
        public static IWordApplication CreateOptimizedWordApplication()
        {
            var app = WordFactory.BlankWorkbook();
            app.Visible = false;
            app.ScreenUpdating = false;
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            return app;
        }

        /// <summary>
        /// 恢复Word应用程序设置
        /// </summary>
        /// <param name="app">Word应用程序</param>
        public static void RestoreWordApplicationSettings(IWordApplication app)
        {
            if (app != null)
            {
                try
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = WdAlertLevel.wdAlertsAll;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"恢复Word应用程序设置时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 处理大型文档
        /// </summary>
        /// <param name="documentPaths">文档路径列表</param>
        /// <param name="batchSize">批处理大小</param>
        /// <param name="processAction">处理操作</param>
        /// <returns>处理结果</returns>
        public static async Task<BatchProcessingResult> ProcessLargeDocumentsAsync(
            List<string> documentPaths,
            int batchSize,
            Action<IWordDocument> processAction)
        {
            var result = new BatchProcessingResult
            {
                TotalDocuments = documentPaths.Count,
                ProcessedDocuments = new List<string>(),
                FailedDocuments = new List<string>()
            };

            for (int i = 0; i < documentPaths.Count; i += batchSize)
            {
                var batch = documentPaths.Skip(i).Take(batchSize).ToList();

                using var app = WordFactory.BlankWorkbook();
                app.Visible = false;
                app.ScreenUpdating = false;
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                foreach (var path in batch)
                {
                    try
                    {
                        var doc = app.Documents.Open(path);
                        try
                        {
                            processAction(doc);
                            result.ProcessedDocuments.Add(path);
                        }
                        finally
                        {
                            doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedDocuments.Add(path);
                        Console.WriteLine($"处理文档 {path} 时出错: {ex.Message}");
                    }
                }

                // 恢复设置
                app.ScreenUpdating = true;

                // 强制垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }

        /// <summary>
        /// 生成性能报告
        /// </summary>
        /// <param name="result">批量处理结果</param>
        /// <param name="processingTime">处理时间</param>
        /// <returns>性能报告</returns>
        public static string GeneratePerformanceReport(BatchProcessingResult result, TimeSpan processingTime)
        {
            var report = new StringBuilder();
            report.AppendLine("=== 性能报告 ===");
            report.AppendLine($"处理时间: {processingTime.TotalSeconds:F2} 秒");
            report.AppendLine($"总文档数: {result.TotalDocuments}");
            report.AppendLine($"成功处理: {result.ProcessedDocuments.Count}");
            report.AppendLine($"处理失败: {result.FailedDocuments.Count}");
            report.AppendLine($"成功率: {(result.TotalDocuments > 0 ? (double)result.ProcessedDocuments.Count / result.TotalDocuments * 100 : 0):F2}%");
            report.AppendLine($"平均每文档处理时间: {(result.ProcessedDocuments.Count > 0 ? processingTime.TotalMilliseconds / result.ProcessedDocuments.Count : 0):F2} 毫秒");

            if (result.FailedDocuments.Any())
            {
                report.AppendLine("\n失败文档列表:");
                foreach (var doc in result.FailedDocuments)
                {
                    report.AppendLine($"  - {doc}");
                }
            }

            report.AppendLine("================");

            return report.ToString();
        }
    }
}