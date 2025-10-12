using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FAQSample
{
    /// <summary>
    /// 常见问题帮助类
    /// </summary>
    public class CommonIssuesHelper
    {
        /// <summary>
        /// 检查Office安装状态
        /// </summary>
        /// <returns>是否安装了Office</returns>
        public static bool IsOfficeInstalled()
        {
            try
            {
                // 尝试创建Word应用程序实例来检查Office是否安装
                using var app = WordFactory.BlankWorkbook();
                return app != null;
            }
            catch (COMException)
            {
                // COM异常通常表示Office未安装或无法访问
                return false;
            }
            catch (Exception)
            {
                // 其他异常也表示可能存在问题
                return false;
            }
        }

        /// <summary>
        /// 安全地释放COM对象
        /// </summary>
        /// <param name="comObject">COM对象</param>
        public static void SafeReleaseComObject(object comObject)
        {
            if (comObject != null)
            {
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"释放COM对象时出错: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 验证文档有效性
        /// </summary>
        /// <param name="filePath">文档路径</param>
        /// <returns>文档是否有效</returns>
        public static bool IsDocumentValid(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            try
            {
                using var app = WordFactory.Open(filePath);
                var doc = app.ActiveDocument;
                // 尝试访问文档属性来验证文档是否有效
                var _ = doc.Paragraphs.Count;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 处理文件不存在异常
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>操作结果</returns>
        public static FileOperationResult HandleFileNotFound(string filePath)
        {
            var result = new FileOperationResult();

            try
            {
                using var app = WordFactory.Open(filePath);
                // 处理文档
                result.Success = true;
                result.Message = "文件操作成功";
            }
            catch (FileNotFoundException ex)
            {
                result.Success = false;
                result.Message = $"文件未找到: {ex.Message}";
                result.ErrorCode = "FILE_NOT_FOUND";
                Console.WriteLine($"文件未找到: {ex.Message}");
            }
            catch (UnauthorizedAccessException ex)
            {
                result.Success = false;
                result.Message = $"访问被拒绝: {ex.Message}";
                result.ErrorCode = "ACCESS_DENIED";
                Console.WriteLine($"访问被拒绝: {ex.Message}");
            }
            catch (COMException ex)
            {
                result.Success = false;
                result.Message = $"COM错误: {ex.Message}";
                result.ErrorCode = $"COM_ERROR_{ex.HResult}";
                Console.WriteLine($"COM错误: {ex.Message}, HRESULT: {ex.HResult}");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"其他错误: {ex.Message}";
                result.ErrorCode = "UNKNOWN_ERROR";
                Console.WriteLine($"其他错误: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 在STA线程中执行Word操作
        /// </summary>
        /// <param name="wordOperation">Word操作</param>
        /// <returns>任务</returns>
        public static async Task ExecuteInStaThreadAsync(Action wordOperation)
        {
            var task = Task.Factory.StartNew(() =>
            {
                Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                wordOperation();
            }, TaskCreationOptions.LongRunning);

            await task;
        }

        /// <summary>
        /// 创建Word实例管理器
        /// </summary>
        public class WordInstanceManager
        {
            private static IWordApplication _sharedInstance;
            private static readonly object _lockObject = new object();

            /// <summary>
            /// 获取共享的Word实例
            /// </summary>
            /// <returns>Word应用程序实例</returns>
            public static IWordApplication GetSharedInstance()
            {
                lock (_lockObject)
                {
                    if (_sharedInstance == null)
                    {
                        _sharedInstance = WordFactory.BlankWorkbook();
                        _sharedInstance.Visible = false;
                    }
                    return _sharedInstance;
                }
            }

            /// <summary>
            /// 释放共享的Word实例
            /// </summary>
            public static void ReleaseSharedInstance()
            {
                lock (_lockObject)
                {
                    if (_sharedInstance != null)
                    {
                        try
                        {
                            _sharedInstance.Quit();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"释放Word实例时出错: {ex.Message}");
                        }
                        finally
                        {
                            _sharedInstance = null;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 批量处理文档
        /// </summary>
        /// <param name="documentPaths">文档路径列表</param>
        /// <returns>处理结果</returns>
        public static async Task<BatchProcessingResult> ProcessDocumentsAsync(List<string> documentPaths)
        {
            var result = new BatchProcessingResult
            {
                TotalDocuments = documentPaths.Count,
                ProcessedDocuments = new List<string>(),
                FailedDocuments = new List<string>()
            };

            const int batchSize = 10;

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
                        // 处理文档（这里简化处理）
                        await Task.Delay(100); // 模拟处理时间
                        doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                        result.ProcessedDocuments.Add(path);
                    }
                    catch (Exception ex)
                    {
                        result.FailedDocuments.Add(path);
                        Console.WriteLine($"处理文档 {path} 时出错: {ex.Message}");
                    }
                }

                // 启用屏幕更新
                app.ScreenUpdating = true;

                // 强制垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }
    }

    /// <summary>
    /// 文件操作结果类
    /// </summary>
    public class FileOperationResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误代码
        /// </summary>
        public string ErrorCode { get; set; }
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