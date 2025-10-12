using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMergeSample
{
    /// <summary>
    /// 邮件合并助手类
    /// </summary>
    public class MailMergeHelper
    {
        private readonly IWordDocument _document;
        private readonly IWordMailMerge _mailMerge;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public MailMergeHelper(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _mailMerge = document.MailMerge ?? throw new ArgumentNullException(nameof(document.MailMerge));
        }

        /// <summary>
        /// 设置邮件合并主文档类型
        /// </summary>
        /// <param name="documentType">主文档类型</param>
        public void SetMainDocumentType(WdMailMergeMainDocType documentType)
        {
            try
            {
                _mailMerge.MainDocumentType = documentType;
                Console.WriteLine($"已设置主文档类型为: {documentType}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"设置主文档类型时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查文档是否为邮件合并文档
        /// </summary>
        /// <returns>是否为邮件合并文档</returns>
        public bool IsMailMergeDocument()
        {
            try
            {
                bool isMailMergeDoc = _mailMerge.MainDocumentType != WdMailMergeMainDocType.wdNotAMergeDocument;
                Console.WriteLine($"是否为邮件合并文档: {isMailMergeDoc}");
                return isMailMergeDoc;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"检查邮件合并文档时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 连接Excel数据源
        /// </summary>
        /// <param name="dataSourcePath">数据源路径</param>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="sqlStatement">SQL语句</param>
        /// <returns>是否连接成功</returns>
        public bool ConnectToExcelDataSource(string dataSourcePath, string connectionString, string sqlStatement)
        {
            try
            {
                _mailMerge.OpenDataSource(
                    Name: dataSourcePath,
                    Connection: connectionString,
                    SQLStatement: sqlStatement
                );

                Console.WriteLine("Excel数据源连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel数据源连接失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 连接CSV文本文件数据源
        /// </summary>
        /// <param name="dataSourcePath">数据源路径</param>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="sqlStatement">SQL语句</param>
        /// <returns>是否连接成功</returns>
        public bool ConnectToCsvDataSource(string dataSourcePath, string connectionString, string sqlStatement)
        {
            try
            {
                _mailMerge.OpenDataSource(
                    Name: dataSourcePath,
                    Connection: connectionString,
                    SQLStatement: sqlStatement
                );

                Console.WriteLine("CSV数据源连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"CSV数据源连接失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加合并字段到指定范围
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="fieldName">字段名称</param>
        /// <returns>是否添加成功</returns>
        public bool AddMergeField(IWordRange range, string fieldName)
        {
            try
            {
                _mailMerge.Fields.Add(range, fieldName);
                Console.WriteLine($"已添加合并字段: {fieldName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加合并字段时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取所有合并字段信息
        /// </summary>
        /// <returns>合并字段信息列表</returns>
        public List<string> GetAllMergeFields()
        {
            var fields = new List<string>();

            try
            {
                Console.WriteLine($"合并字段数量: {_mailMerge.Fields.Count}");
                for (int i = 1; i <= _mailMerge.Fields.Count; i++)
                {
                    string fieldCode = _mailMerge.Fields.Item(i).Code;
                    fields.Add(fieldCode);
                    Console.WriteLine($"字段 {i}: {fieldCode}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取合并字段信息时出错: {ex.Message}");
            }

            return fields;
        }

        /// <summary>
        /// 执行邮件合并
        /// </summary>
        /// <param name="destination">合并结果目标</param>
        /// <param name="pause">是否暂停</param>
        /// <returns>是否执行成功</returns>
        public bool ExecuteMerge(WdMailMergeDestination destination = WdMailMergeDestination.wdSendToNewDocument, bool pause = false)
        {
            try
            {
                _mailMerge.Destination = destination;
                _mailMerge.Execute(Pause: pause);

                Console.WriteLine("邮件合并执行完成");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"邮件合并执行失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 执行邮件合并到打印机
        /// </summary>
        /// <param name="pause">是否暂停</param>
        /// <returns>是否执行成功</returns>
        public bool ExecuteMergeToPrinter(bool pause = false)
        {
            return ExecuteMerge(WdMailMergeDestination.wdSendToPrinter, pause);
        }

        /// <summary>
        /// 执行邮件合并到新文档
        /// </summary>
        /// <param name="pause">是否暂停</param>
        /// <returns>是否执行成功</returns>
        public bool ExecuteMergeToNewDocument(bool pause = false)
        {
            return ExecuteMerge(WdMailMergeDestination.wdSendToNewDocument, pause);
        }

        /// <summary>
        /// 执行邮件合并到电子邮件
        /// </summary>
        /// <param name="pause">是否暂停</param>
        /// <returns>是否执行成功</returns>
        public bool ExecuteMergeToEmail(bool pause = false)
        {
            return ExecuteMerge(WdMailMergeDestination.wdSendToEmail, pause);
        }

        /// <summary>
        /// 添加条件合并字段
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="conditionText">条件文本</param>
        /// <returns>是否添加成功</returns>
        public bool AddConditionalField(IWordRange range, string conditionText)
        {
            try
            {
                range.Text = conditionText;
                Console.WriteLine($"已添加条件字段: {conditionText}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加条件字段时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加计算字段
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="calculationText">计算文本</param>
        /// <returns>是否添加成功</returns>
        public bool AddCalculationField(IWordRange range, string calculationText)
        {
            try
            {
                range.Text = calculationText;
                Console.WriteLine($"已添加计算字段: {calculationText}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加计算字段时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 添加日期字段
        /// </summary>
        /// <param name="range">文档范围</param>
        /// <param name="dateFormat">日期格式</param>
        /// <returns>是否添加成功</returns>
        public bool AddDateField(IWordRange range, string dateFormat = "yyyy年MM月dd日")
        {
            try
            {
                range.Text = $"{{ DATE \\@ \"{dateFormat}\" }}";
                Console.WriteLine($"已添加日期字段，格式: {dateFormat}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"添加日期字段时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 更新所有字段
        /// </summary>
        /// <returns>是否更新成功</returns>
        public bool UpdateAllFields()
        {
            try
            {
                _document.Range().Fields.Update();
                Console.WriteLine("所有字段已更新");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"更新字段时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取邮件合并状态信息
        /// </summary>
        /// <returns>状态信息</returns>
        public MailMergeStatus GetMailMergeStatus()
        {
            var status = new MailMergeStatus();

            try
            {
                status.IsMailMergeDocument = IsMailMergeDocument();
                status.MainDocumentType = _mailMerge.MainDocumentType;
                status.FieldCount = _mailMerge.Fields.Count;
                status.DataSourceName = _mailMerge.DataSource.Name;
                status.DataSourceRecordCount = _mailMerge.DataSource.RecordCount;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取邮件合并状态时出错: {ex.Message}");
                status.ErrorMessage = ex.Message;
            }

            return status;
        }
    }

    /// <summary>
    /// 邮件合并状态信息类
    /// </summary>
    public class MailMergeStatus
    {
        /// <summary>
        /// 是否为邮件合并文档
        /// </summary>
        public bool IsMailMergeDocument { get; set; }

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
        /// 数据源记录数
        /// </summary>
        public int DataSourceRecordCount { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成状态报告
        /// </summary>
        /// <returns>状态报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"获取状态失败: {ErrorMessage}";
            }

            return $"邮件合并状态报告:\n" +
                   $"  是否为邮件合并文档: {IsMailMergeDocument}\n" +
                   $"  主文档类型: {MainDocumentType}\n" +
                   $"  合并字段数量: {FieldCount}\n" +
                   $"  数据源名称: {DataSourceName}\n" +
                   $"  数据源记录数: {DataSourceRecordCount}";
        }
    }
}