//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelWorkbookConnection : IExcelWorkbookConnection
{
    private MsExcel.WorkbookConnection _connection;
    private bool _disposedValue;

    public IExcelApplication Application
    {
        get
        {
            var application = _connection?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public string Name
    {
        get => _connection.Name;
        set => _connection.Name = value;
    }

    public string Description
    {
        get => _connection.Description;
        set => _connection.Description = value;
    }

    public IExcelRanges Ranges => new ExcelRanges(_connection.Ranges);


    public XlConnectionType Type => (XlConnectionType)(int)_connection.Type;

    public IExcelConnections Parent => new ExcelConnections(_connection.Parent as MsExcel.Connections);

    public IExcelOLEDBConnection OLEDBConnection => new ExcelOLEDBConnection(_connection.OLEDBConnection);

    public IExcelODBCConnection ODBCConnection => new ExcelODBCConnection(_connection.ODBCConnection);

    public IExcelModelConnection ModelConnection => new ExcelModelConnection(_connection.ModelConnection);

    public IExcelWorksheetDataConnection WorksheetDataConnection => new ExcelWorksheetDataConnection(_connection.WorksheetDataConnection);

    public IExcelTextConnection TextConnection => new ExcelTextConnection(_connection.TextConnection);

    public IExcelDataFeedConnection DataFeedConnection => new ExcelDataFeedConnection(_connection.DataFeedConnection);

    internal ExcelWorkbookConnection(MsExcel.WorkbookConnection connection)
    {
        _connection = connection ?? throw new ArgumentNullException(nameof(connection));
        _disposedValue = false;
    }

    public void Refresh()
    {
        try
        {
            _connection.Refresh();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法刷新连接: {Name}", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _connection.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法删除连接: {Name}", ex);
        }
    }

    public bool TestConnection()
    {
        try
        {
            _connection.Refresh();
            return true;
        }
        catch (COMException)
        {
            return false;
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _connection != null)
        {
            try
            {
                Marshal.ReleaseComObject(_connection);
            }
            catch { }
            _connection = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}