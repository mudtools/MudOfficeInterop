//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelODBCConnection : IExcelODBCConnection
{
    private MsExcel.ODBCConnection _odbcConnection;
    private bool _disposedValue;

    public IExcelApplication Application
    {
        get
        {
            var application = _odbcConnection?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public string CommandText
    {
        get => _odbcConnection.CommandText?.ToString();
        set => _odbcConnection.CommandText = value;
    }

    public XlCmdType CommandType
    {
        get => (XlCmdType)(int)_odbcConnection.CommandType;
        set => _odbcConnection.CommandType = (MsExcel.XlCmdType)value;
    }

    public string Connection
    {
        get => _odbcConnection.Connection.ToString();
        set => _odbcConnection.Connection = value;
    }


    public bool BackgroundQuery
    {
        get => _odbcConnection.BackgroundQuery;
        set => _odbcConnection.BackgroundQuery = value;
    }

    public IExcelWorkbookConnection Parent => new ExcelWorkbookConnection(_odbcConnection.Parent as MsExcel.WorkbookConnection);

    public bool EnableRefresh
    {
        get => _odbcConnection.EnableRefresh;
        set => _odbcConnection.EnableRefresh = value;
    }

    public bool RefreshOnFileOpen
    {
        get => _odbcConnection.RefreshOnFileOpen;
        set => _odbcConnection.RefreshOnFileOpen = value;
    }

    public bool SavePassword
    {
        get => _odbcConnection.SavePassword;
        set => _odbcConnection.SavePassword = value;
    }

    public string SourceDataFile
    {
        get => _odbcConnection.SourceDataFile;
        set => _odbcConnection.SourceDataFile = value;
    }

    public string SourceConnectionFile
    {
        get => _odbcConnection.SourceConnectionFile;
        set => _odbcConnection.SourceConnectionFile = value;
    }

    public bool AlwaysUseConnectionFile
    {
        get => _odbcConnection.AlwaysUseConnectionFile;
        set => _odbcConnection.AlwaysUseConnectionFile = value;
    }

    internal ExcelODBCConnection(MsExcel.ODBCConnection odbcConnection)
    {
        _odbcConnection = odbcConnection ?? throw new ArgumentNullException(nameof(odbcConnection));
        _disposedValue = false;
    }

    public void Refresh()
    {
        try
        {
            _odbcConnection.Refresh();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法刷新ODBC连接。", ex);
        }
    }

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

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _odbcConnection != null)
        {
            try
            {
                Marshal.ReleaseComObject(_odbcConnection);
            }
            catch { }
            _odbcConnection = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
