using MudTools.OfficeInterop.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMergeSample
{
    /// <summary>
    /// 数据源管理器类
    /// </summary>
    public class DataSourceManager
    {
        private readonly IWordDocument _document;
        private readonly IWordMailMerge _mailMerge;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="document">Word文档对象</param>
        public DataSourceManager(IWordDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _mailMerge = document.MailMerge ?? throw new ArgumentNullException(nameof(document.MailMerge));
        }

        /// <summary>
        /// 创建示例Excel数据源
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="data">数据</param>
        /// <returns>是否创建成功</returns>
        public bool CreateSampleExcelDataSource(string filePath, List<Dictionary<string, object>> data)
        {
            try
            {
                // 注意：在实际应用中，这里需要使用Excel.Interop或其他库来创建Excel文件
                // 为了简化示例，我们只是创建一个空文件来模拟
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                File.WriteAllText(filePath, "这是一个模拟的Excel数据源文件");

                Console.WriteLine($"示例Excel数据源已创建: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建示例Excel数据源时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建示例CSV数据源
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="headers">表头</param>
        /// <param name="data">数据</param>
        /// <returns>是否创建成功</returns>
        public bool CreateSampleCsvDataSource(string filePath, List<string> headers, List<List<string>> data)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                var lines = new List<string>();

                // 添加表头
                lines.Add(string.Join(",", headers));

                // 添加数据行
                foreach (var row in data)
                {
                    lines.Add(string.Join(",", row.Select(field => $"\"{field}\"")));
                }

                File.WriteAllLines(filePath, lines, Encoding.UTF8);

                Console.WriteLine($"示例CSV数据源已创建: {filePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建示例CSV数据源时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 连接Excel数据源
        /// </summary>
        /// <param name="dataSourcePath">数据源路径</param>
        /// <param name="worksheetName">工作表名称</param>
        /// <returns>是否连接成功</returns>
        public bool ConnectToExcelDataSource(string dataSourcePath, string worksheetName = "Sheet1")
        {
            try
            {
                string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dataSourcePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";
                string sqlStatement = $"SELECT * FROM [{worksheetName}$]";

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
        /// 连接CSV数据源
        /// </summary>
        /// <param name="dataSourcePath">数据源路径</param>
        /// <returns>是否连接成功</returns>
        public bool ConnectToCsvDataSource(string dataSourcePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(dataSourcePath);
                string fileName = Path.GetFileName(dataSourcePath);
                
                string connectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={directory};Extended Properties=\"text;HDR=YES;FMT=Delimited\"";
                string sqlStatement = $"SELECT * FROM [{fileName}]";

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
        /// 连接Access数据库数据源
        /// </summary>
        /// <param name="dataSourcePath">数据源路径</param>
        /// <param name="tableName">表名</param>
        /// <returns>是否连接成功</returns>
        public bool ConnectToAccessDataSource(string dataSourcePath, string tableName)
        {
            try
            {
                string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dataSourcePath};Persist Security Info=False;";
                string sqlStatement = $"SELECT * FROM [{tableName}]";

                _mailMerge.OpenDataSource(
                    Name: dataSourcePath,
                    Connection: connectionString,
                    SQLStatement: sqlStatement
                );

                Console.WriteLine("Access数据库数据源连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Access数据库数据源连接失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 连接SQL Server数据源
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="sqlStatement">SQL语句</param>
        /// <returns>是否连接成功</returns>
        public bool ConnectToSqlServerDataSource(string connectionString, string sqlStatement)
        {
            try
            {
                _mailMerge.OpenDataSource(
                    Name: "SQL Server",
                    Connection: connectionString,
                    SQLStatement: sqlStatement
                );

                Console.WriteLine("SQL Server数据源连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SQL Server数据源连接失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取数据源信息
        /// </summary>
        /// <returns>数据源信息</returns>
        public DataSourceInfo GetDataSourceInfo()
        {
            var info = new DataSourceInfo();

            try
            {
                info.Name = _mailMerge.DataSource.Name;
                info.HeaderSource = _mailMerge.DataSource.HeaderSource;
                info.RecordCount = _mailMerge.DataSource.RecordCount;
                info.FieldNames = new List<string>();

                // 获取字段名称
                for (int i = 1; i <= _mailMerge.DataSource.FieldNames.Count; i++)
                {
                    info.FieldNames.Add(_mailMerge.DataSource.FieldNames.Item(i));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取数据源信息时出错: {ex.Message}");
                info.ErrorMessage = ex.Message;
            }

            return info;
        }

        /// <summary>
        /// 创建客户数据示例
        /// </summary>
        /// <returns>客户数据列表</returns>
        public List<Dictionary<string, object>> CreateSampleCustomerData()
        {
            var data = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    {"客户姓名", "张三"},
                    {"地址", "北京市朝阳区某某街道123号"},
                    {"客户编号", "C12345"},
                    {"账户余额", 10000},
                    {"信用等级", "A"},
                    {"性别", "男"}
                },
                new Dictionary<string, object>
                {
                    {"客户姓名", "李四"},
                    {"地址", "上海市浦东新区某某路456号"},
                    {"客户编号", "C67890"},
                    {"账户余额", 25000},
                    {"信用等级", "AA"},
                    {"性别", "女"}
                },
                new Dictionary<string, object>
                {
                    {"客户姓名", "王五"},
                    {"地址", "广州市天河区某某大道789号"},
                    {"客户编号", "C24680"},
                    {"账户余额", 15000},
                    {"信用等级", "A"},
                    {"性别", "男"}
                }
            };

            return data;
        }

        /// <summary>
        /// 创建产品数据示例
        /// </summary>
        /// <returns>产品数据列表</returns>
        public List<Dictionary<string, object>> CreateSampleProductData()
        {
            var data = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    {"产品名称", "产品A"},
                    {"产品描述", "这是一款高质量的产品A"},
                    {"价格", 299.99}
                },
                new Dictionary<string, object>
                {
                    {"产品名称", "产品B"},
                    {"产品描述", "这是一款功能强大的产品B"},
                    {"价格", 399.99}
                },
                new Dictionary<string, object>
                {
                    {"产品名称", "产品C"},
                    {"产品描述", "这是一款性价比高的产品C"},
                    {"价格", 199.99}
                }
            };

            return data;
        }

        /// <summary>
        /// 创建收件人数据示例
        /// </summary>
        /// <returns>收件人数据列表</returns>
        public List<Dictionary<string, object>> CreateSampleRecipientData()
        {
            var data = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    {"姓名", "张三"},
                    {"地址", "北京市朝阳区某某街道123号"},
                    {"城市", "北京"},
                    {"邮编", "100000"},
                    {"收件人姓名", "张三"},
                    {"收件人地址", "北京市朝阳区某某街道123号"},
                    {"收件人城市", "北京"},
                    {"收件人邮编", "100000"},
                    {"发件人姓名", "ABC有限公司"},
                    {"发件人地址", "某某市某某区某某路123号"},
                    {"发件人城市", "某某市"},
                    {"发件人邮编", "123456"}
                },
                new Dictionary<string, object>
                {
                    {"姓名", "李四"},
                    {"地址", "上海市浦东新区某某路456号"},
                    {"城市", "上海"},
                    {"邮编", "200000"},
                    {"收件人姓名", "李四"},
                    {"收件人地址", "上海市浦东新区某某路456号"},
                    {"收件人城市", "上海"},
                    {"收件人邮编", "200000"},
                    {"发件人姓名", "ABC有限公司"},
                    {"发件人地址", "某某市某某区某某路123号"},
                    {"发件人城市", "某某市"},
                    {"发件人邮编", "123456"}
                }
            };

            return data;
        }
    }

    /// <summary>
    /// 数据源信息类
    /// </summary>
    public class DataSourceInfo
    {
        /// <summary>
        /// 数据源名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 头部源
        /// </summary>
        public string HeaderSource { get; set; }

        /// <summary>
        /// 记录数
        /// </summary>
        public int RecordCount { get; set; }

        /// <summary>
        /// 字段名称列表
        /// </summary>
        public List<string> FieldNames { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 生成数据源信息报告
        /// </summary>
        /// <returns>信息报告</returns>
        public string GenerateReport()
        {
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                return $"获取数据源信息失败: {ErrorMessage}";
            }

            var fieldNames = FieldNames != null ? string.Join(", ", FieldNames) : "无字段信息";

            return $"数据源信息报告:\n" +
                   $"  数据源名称: {Name}\n" +
                   $"  头部源: {HeaderSource}\n" +
                   $"  记录数: {RecordCount}\n" +
                   $"  字段名称: {fieldNames}";
        }
    }
}