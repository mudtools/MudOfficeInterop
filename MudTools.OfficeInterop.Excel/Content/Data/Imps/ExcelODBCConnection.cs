//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal partial class ExcelODBCConnection : IExcelODBCConnection
{

    public bool TestConnection()
    {
        try
        {
            // 尝试执行一个简单的查询来测试连接
            var originalCommandText = CommandText;
            var originalCommandType = CommandType;

            CommandText = "SELECT 1";
            CommandType = XlCmdType.xlCmdSql;

            Refresh();

            CommandText = originalCommandText;
            CommandType = originalCommandType;

            return true;
        }
        catch (COMException)
        {
            return false;
        }
    }

    public int ExecuteCommand(string sql)
    {
        if (string.IsNullOrEmpty(sql))
            throw new ArgumentException("SQL命令不能为空。", nameof(sql));

        try
        {
            var originalCommandText = CommandText;
            var originalCommandType = CommandType;

            CommandText = sql;
            CommandType = XlCmdType.xlCmdSql;

            Refresh(); // 执行命令

            CommandText = originalCommandText;
            CommandType = originalCommandType;

            return 0; // 伪代码返回值
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法执行SQL命令: {sql}", ex);
        }
    }

}
