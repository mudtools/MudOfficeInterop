//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word;

namespace MailMergeSample
{
    /// <summary>
    /// 邮件合并执行器类
    /// </summary>
    public class MailMergeExecutor
    {
        private readonly IWordApplication _application;
        private readonly IWordDocument _document;
        private readonly IWordMailMerge _mailMerge;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="application">Word应用程序对象</param>
        /// <param name="document">Word文档对象</param>
        public MailMergeExecutor(IWordApplication application, IWordDocument document)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _mailMerge = document.MailMerge ?? throw new ArgumentNullException(nameof(document.MailMerge));
        }

        /// <summary>
        /// 执行邮件合并到新文档
        /// </summary>
        /// <param name="outputPath">输出路径</param>
        /// <param name="pause">是否暂停</param>
        /// <returns>执行结果</returns>
        public MailMergeExecutionResult ExecuteToNewDocument(string outputPath = null, bool pause = false)
        {
            var result = new MailMergeExecutionResult
            {
                ExecutionType = "新文档",
                StartTime = DateTime.Now
            };

            try
            {
                // 设置合并目标
                _mailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;

                // 执行合并
                _mailMerge.Execute(pause: pause);

                result.EndTime = DateTime.Now;
                result.Success = true;

                // 保存结果文档
                if (!string.IsNullOrEmpty(outputPath))
                {
                    var resultDoc = _application.ActiveDocument;
                    resultDoc.Save(outputPath);
                    result.OutputPath = outputPath;
                    Console.WriteLine($"邮件合并结果已保存: {outputPath}");
                }

                Console.WriteLine("邮件合并执行完成，结果发送到新文档");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Console.WriteLine($"邮件合并执行失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 执行邮件合并到打印机
        /// </summary>
        /// <param name="pause">是否暂停</param>
        /// <returns>执行结果</returns>
        public MailMergeExecutionResult ExecuteToPrinter(bool pause = false)
        {
            var result = new MailMergeExecutionResult
            {
                ExecutionType = "打印机",
                StartTime = DateTime.Now
            };

            try
            {
                // 设置合并目标
                _mailMerge.Destination = WdMailMergeDestination.wdSendToPrinter;

                // 执行合并
                _mailMerge.Execute(pause: pause);

                result.EndTime = DateTime.Now;
                result.Success = true;

                Console.WriteLine("邮件合并执行完成，结果发送到打印机");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Console.WriteLine($"邮件合并执行失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 执行邮件合并到电子邮件
        /// </summary>
        /// <param name="pause">是否暂停</param>
        /// <returns>执行结果</returns>
        public MailMergeExecutionResult ExecuteToEmail(bool pause = false)
        {
            var result = new MailMergeExecutionResult
            {
                ExecutionType = "电子邮件",
                StartTime = DateTime.Now
            };

            try
            {
                // 设置合并目标
                _mailMerge.Destination = WdMailMergeDestination.wdSendToEmail;

                // 执行合并
                _mailMerge.Execute(pause: pause);

                result.EndTime = DateTime.Now;
                result.Success = true;

                Console.WriteLine("邮件合并执行完成，结果发送到电子邮件");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Console.WriteLine($"邮件合并执行失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 执行邮件合并到Fax
        /// </summary>
        /// <param name="pause">是否暂停</param>
        /// <returns>执行结果</returns>
        public MailMergeExecutionResult ExecuteToFax(bool pause = false)
        {
            var result = new MailMergeExecutionResult
            {
                ExecutionType = "传真",
                StartTime = DateTime.Now
            };

            try
            {
                // 设置合并目标
                _mailMerge.Destination = WdMailMergeDestination.wdSendToFax;

                // 执行合并
                _mailMerge.Execute(pause: pause);

                result.EndTime = DateTime.Now;
                result.Success = true;

                Console.WriteLine("邮件合并执行完成，结果发送到传真");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Console.WriteLine($"邮件合并执行失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 执行邮件合并的第一条记录
        /// </summary>
        /// <returns>是否执行成功</returns>
        public bool ExecuteFirstRecord()
        {
            try
            {
                _mailMerge.DataSource.FirstRecord = 1;
                _mailMerge.DataSource.LastRecord = 1;
                _mailMerge.Execute(pause: false);

                Console.WriteLine("第一条记录邮件合并执行完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行第一条记录邮件合并时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 执行邮件合并的指定范围记录
        /// </summary>
        /// <param name="firstRecord">起始记录</param>
        /// <param name="lastRecord">结束记录</param>
        /// <returns>是否执行成功</returns>
        public bool ExecuteRecordRange(int firstRecord, int lastRecord)
        {
            try
            {
                _mailMerge.DataSource.FirstRecord = firstRecord;
                _mailMerge.DataSource.LastRecord = lastRecord;
                _mailMerge.Execute(pause: false);

                Console.WriteLine($"记录范围({firstRecord}-{lastRecord})邮件合并执行完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行记录范围邮件合并时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 使用示例数据执行邮件合并
        /// </summary>
        /// <param name="sampleData">示例数据</param>
        /// <param name="outputPath">输出路径</param>
        /// <returns>执行结果</returns>
        public MailMergeExecutionResult ExecuteWithSampleData(List<Dictionary<string, object>> sampleData, string outputPath = null)
        {
            var result = new MailMergeExecutionResult
            {
                ExecutionType = "示例数据",
                StartTime = DateTime.Now
            };

            try
            {
                // 添加示例数据到数据源
                foreach (var data in sampleData)
                {
                    foreach (var field in data)
                    {
                        // 注意：在实际应用中，这里需要通过特定方式添加数据
                        // 由于Word.Interop的限制，我们仅做示意处理
                        Console.WriteLine($"添加示例数据: {field.Key} = {field.Value}");
                    }
                }

                // 设置合并目标
                _mailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;

                // 执行合并
                _mailMerge.Execute(pause: false);

                result.EndTime = DateTime.Now;
                result.Success = true;

                // 保存结果文档
                if (!string.IsNullOrEmpty(outputPath))
                {
                    var resultDoc = _application.ActiveDocument;
                    resultDoc.Save(outputPath);
                    result.OutputPath = outputPath;
                    Console.WriteLine($"邮件合并结果已保存: {outputPath}");
                }

                Console.WriteLine("使用示例数据的邮件合并执行完成");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                result.EndTime = DateTime.Now;
                Console.WriteLine($"使用示例数据的邮件合并执行失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 获取邮件合并统计信息
        /// </summary>
        /// <returns>统计信息</returns>
        public MailMergeStatistics GetMergeStatistics()
        {
            var stats = new MailMergeStatistics();

            try
            {
                stats.MainDocumentType = _mailMerge.MainDocumentType;
                stats.FieldCount = _mailMerge.Fields.Count;
                stats.DataSourceName = _mailMerge.DataSource.Name;
                stats.RecordCount = _mailMerge.DataSource.RecordCount;
                stats.FirstRecord = _mailMerge.DataSource.FirstRecord;
                stats.LastRecord = _mailMerge.DataSource.LastRecord;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取邮件合并统计信息时出错: {ex.Message}");
                stats.ErrorMessage = ex.Message;
            }

            return stats;
        }

        /// <summary>
        /// 预览邮件合并结果
        /// </summary>
        /// <returns>是否预览成功</returns>
        public bool PreviewMergeResult()
        {
            try
            {
                // 设置要预览的记录
                _mailMerge.DataSource.ActiveRecord = WdMailMergeActiveRecord.wdFirstRecord;

                // 更新文档中的合并字段
                _document.Range().Fields.Update();

                Console.WriteLine($"已预览第 1 条记录的合并结果");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"预览邮件合并结果时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 检查邮件合并设置
        /// </summary>
        /// <returns>检查结果</returns>
        public MailMergeCheckResult CheckMergeSettings()
        {
            var checkResult = new MailMergeCheckResult();

            try
            {
                // 检查是否为邮件合并文档
                checkResult.IsMailMergeDocument = _mailMerge.MainDocumentType != WdMailMergeMainDocType.wdNotAMergeDocument;

                // 检查是否有合并字段
                checkResult.HasMergeFields = _mailMerge.Fields.Count > 0;

                // 检查是否有数据源
                checkResult.HasDataSource = !string.IsNullOrEmpty(_mailMerge.DataSource.Name);

                // 检查数据源记录数
                if (checkResult.HasDataSource)
                {
                    checkResult.RecordCount = _mailMerge.DataSource.RecordCount;
                }

                checkResult.IsValid = checkResult.IsMailMergeDocument &&
                                     checkResult.HasMergeFields &&
                                     checkResult.HasDataSource &&
                                     checkResult.RecordCount > 0;

                Console.WriteLine("邮件合并设置检查完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"检查邮件合并设置时出错: {ex.Message}");
                checkResult.ErrorMessage = ex.Message;
            }

            return checkResult;
        }
    }

    /// <summary>
    /// 邮件合并执行结果类
    /// </summary>
    public class MailMergeExecutionResult
    {
        /// <summary>
        /// 执行类型
        /// </summary>
        public string ExecutionType { get; set; }

        /// <summary>
        /// 是否执行成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 执行时长（秒）
        /// </summary>
        public double Duration => (EndTime - StartTime).TotalSeconds;

        /// <summary>
        /// 输出路径
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成执行报告
        /// </summary>
        /// <returns>执行报告</returns>
        public string GenerateReport()
        {
            if (!Success)
            {
                return $"邮件合并执行失败: {ErrorMessage}";
            }

            return $"邮件合并执行报告:\n" +
                   $"  执行类型: {ExecutionType}\n" +
                   $"  执行状态: 成功\n" +
                   $"  开始时间: {StartTime:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  结束时间: {EndTime:yyyy-MM-dd HH:mm:ss}\n" +
                   $"  执行时长: {Duration:F2} 秒\n" +
                   $"  输出路径: {OutputPath}";
        }
    }

    /// <summary>
    /// 邮件合并统计信息类
    /// </summary>
    public class MailMergeStatistics
    {
        /// <summary>
        /// 主文档类型
        /// </summary>
        public WdMailMergeMainDocType MainDocumentType { get; set; }

        /// <summary>
        /// 字段数量
        /// </summary>
        public int FieldCount { get; set; }

        /// <summary>
        /// 数据源名称
        /// </summary>
        public string DataSourceName { get; set; }

        /// <summary>
        /// 记录数
        /// </summary>
        public int RecordCount { get; set; }

        /// <summary>
        /// 起始记录
        /// </summary>
        public int FirstRecord { get; set; }

        /// <summary>
        /// 结束记录
        /// </summary>
        public int LastRecord { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成统计报告
        /// </summary>
        /// <returns>统计报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"获取统计信息失败: {ErrorMessage}";
            }

            return $"邮件合并统计信息:\n" +
                   $"  主文档类型: {MainDocumentType}\n" +
                   $"  合并字段数量: {FieldCount}\n" +
                   $"  数据源名称: {DataSourceName}\n" +
                   $"  数据源记录数: {RecordCount}\n" +
                   $"  起始记录: {FirstRecord}\n" +
                   $"  结束记录: {LastRecord}";
        }
    }

    /// <summary>
    /// 邮件合并检查结果类
    /// </summary>
    public class MailMergeCheckResult
    {
        /// <summary>
        /// 是否为邮件合并文档
        /// </summary>
        public bool IsMailMergeDocument { get; set; }

        /// <summary>
        /// 是否有合并字段
        /// </summary>
        public bool HasMergeFields { get; set; }

        /// <summary>
        /// 是否有数据源
        /// </summary>
        public bool HasDataSource { get; set; }

        /// <summary>
        /// 记录数
        /// </summary>
        public int RecordCount { get; set; }

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成检查报告
        /// </summary>
        /// <returns>检查报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"检查失败: {ErrorMessage}";
            }

            return $"邮件合并设置检查报告:\n" +
                   $"  是否为邮件合并文档: {IsMailMergeDocument}\n" +
                   $"  是否有合并字段: {HasMergeFields}\n" +
                   $"  是否有数据源: {HasDataSource}\n" +
                   $"  数据源记录数: {RecordCount}\n" +
                   $"  设置是否有效: {IsValid}";
        }
    }
}