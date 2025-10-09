//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// Excel ChartGroups 集合对象的二次封装实现类
/// 实现 IExcelChartGroups 接口
/// </summary>
internal class ExcelChartGroups : IExcelChartGroups
{
    private MsExcel.ChartGroups _chartGroups;
    private bool _disposedValue = false;

    internal ExcelChartGroups(MsExcel.ChartGroups chartGroups)
    {
        _chartGroups = chartGroups ?? throw new ArgumentNullException(nameof(chartGroups));
    }

    #region 基础属性
    public int Count => _chartGroups.Count;

    public IExcelChartGroup? this[int index]
    {
        get
        {
            return _chartGroups != null ? new ExcelChartGroup(_chartGroups.Item(index)) : null;
        }
    }

    public IExcelChartGroup? this[string name]
    {
        get
        {
            return _chartGroups != null ? new ExcelChartGroup(_chartGroups.Item(name)) : null;
        }
    }

    public object? Parent => _chartGroups.Parent;

    public IExcelApplication? Application => new ExcelApplication(_chartGroups.Application);
    #endregion

    #region IEnumerable<IExcelChartGroup> Support
    public IEnumerator<IExcelChartGroup> GetEnumerator()
    {
        for (int i = 1; i <= _chartGroups.Count; i++)
        {
            yield return new ExcelChartGroup(_chartGroups.Item(i));
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (_chartGroups != null)
            {
                Marshal.ReleaseComObject(_chartGroups);
                _chartGroups = null;
            }
            _disposedValue = true;
        }
    }

    ~ExcelChartGroups()
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