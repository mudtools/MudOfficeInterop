//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

internal class ExcelTextConnection : IExcelTextConnection
{
    private MsExcel.TextConnection _textConnection;
    private bool _disposedValue;


    public IExcelWorkbookConnection Parent => new ExcelWorkbookConnection(_textConnection.Parent as MsExcel.WorkbookConnection);

    public IExcelApplication Application
    {
        get
        {
            var application = _textConnection?.Application;
            return application != null ? new ExcelApplication(application) : null;
        }
    }

    public int TextFileStartRow
    {
        get => _textConnection.TextFileStartRow;
        set => _textConnection.TextFileStartRow = value;
    }

    public XlTextParsingType TextFileParseType
    {
        get => (XlTextParsingType)(int)_textConnection.TextFileParseType;
        set => _textConnection.TextFileParseType = (MsExcel.XlTextParsingType)value;
    }

    public XlTextQualifier TextFileTextQualifier
    {
        get => (XlTextQualifier)(int)_textConnection.TextFileTextQualifier;
        set => _textConnection.TextFileTextQualifier = (MsExcel.XlTextQualifier)value;
    }

    public int[] TextFileFixedColumnWidths
    {
        get => _textConnection.TextFileFixedColumnWidths as int[];
        set => _textConnection.TextFileFixedColumnWidths = value;
    }

    public XlPlatform TextFilePlatform
    {
        get => (XlPlatform)_textConnection.TextFilePlatform;
        set => _textConnection.TextFilePlatform = (MsExcel.XlPlatform)value;
    }

    internal ExcelTextConnection(MsExcel.TextConnection textConnection)
    {
        _textConnection = textConnection ?? throw new ArgumentNullException(nameof(textConnection));
        _disposedValue = false;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _textConnection != null)
        {
            try
            {
                Marshal.ReleaseComObject(_textConnection);
            }
            catch { }
            _textConnection = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}