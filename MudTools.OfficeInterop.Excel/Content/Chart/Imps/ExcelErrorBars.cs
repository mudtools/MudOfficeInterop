//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ErrorBars 对象的二次封装实现类
/// 实现 IExcelErrorBars 接口
/// </summary>
internal class ExcelErrorBars : IExcelErrorBars
{
    private MsExcel.ErrorBars? _errorBars;
    private bool _disposedValue = false;

    internal ExcelErrorBars(MsExcel.ErrorBars errorBars)
    {
        _errorBars = errorBars ?? throw new ArgumentNullException(nameof(errorBars));
    }

    #region 基础属性
    public string Name => _errorBars?.Name ?? String.Empty;

    public object? Parent => _errorBars?.Parent;

    public IExcelApplication? Application => _errorBars != null ? new ExcelApplication(_errorBars.Application) : null;
    #endregion

    #region 格式设置
    public IExcelBorder? Border => _errorBars != null ? new ExcelBorder(_errorBars.Border) : null;

    public IExcelChartFormat? Format => _errorBars != null ? new ExcelChartFormat(_errorBars.Format) : null;

    public XlEndStyleCap EndStyle
    {
        get => _errorBars?.EndStyle.EnumConvert(XlEndStyleCap.xlNoCap) ?? XlEndStyleCap.xlNoCap;
        set
        {
            if (_errorBars != null)
                _errorBars.EndStyle = value.EnumConvert(MsExcel.XlEndStyleCap.xlNoCap);
        }
    }
    #endregion

    #region 操作方法
    public void Select()
    {
        _errorBars?.Select();
    }

    public void Delete()
    {
        _errorBars?.Delete();
    }

    public void ClearFormats()
    {
        _errorBars?.ClearFormats();
    }
    #endregion   

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_errorBars != null)
                Marshal.ReleaseComObject(_errorBars);
            _errorBars = null;
        }
        _disposedValue = true;

    }

    ~ExcelErrorBars()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
