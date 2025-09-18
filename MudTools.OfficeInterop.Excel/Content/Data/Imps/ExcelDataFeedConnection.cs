//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelDataFeedConnection : IExcelDataFeedConnection
{
    private MsExcel.DataFeedConnection _dataFeedConnection;
    private bool _disposedValue;

    public IExcelApplication Application
    {
        get
        {
            var application = _dataFeedConnection?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public string CommandText
    {
        get => _dataFeedConnection.CommandText?.ToString();
        set => _dataFeedConnection.CommandText = value;
    }

    public XlCmdType CommandType
    {
        get => (XlCmdType)(int)_dataFeedConnection.CommandType;
        set => _dataFeedConnection.CommandType = (MsExcel.XlCmdType)value;
    }

    public IExcelWorkbookConnection Parent => new ExcelWorkbookConnection(_dataFeedConnection.Parent as MsExcel.WorkbookConnection);

    public bool RefreshOnFileOpen
    {
        get => _dataFeedConnection.RefreshOnFileOpen;
        set => _dataFeedConnection.RefreshOnFileOpen = value;
    }

    public bool SavePassword
    {
        get => _dataFeedConnection.SavePassword;
        set => _dataFeedConnection.SavePassword = value;
    }

    internal ExcelDataFeedConnection(MsExcel.DataFeedConnection dataFeedConnection)
    {
        _dataFeedConnection = dataFeedConnection ?? throw new ArgumentNullException(nameof(dataFeedConnection));
        _disposedValue = false;
    }

    public void Refresh()
    {
        try
        {
            _dataFeedConnection.Refresh();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法刷新数据源连接。", ex);
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dataFeedConnection != null)
        {
            try
            {
                Marshal.ReleaseComObject(_dataFeedConnection);
            }
            catch { }
            _dataFeedConnection = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}