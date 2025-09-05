//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel PivotCache 对象的二次封装实现类
/// 实现 IExcelPivotCache 接口
/// </summary>
internal class ExcelPivotCache : IExcelPivotCache
{
    internal MsExcel.PivotCache _pivotCache;
    private bool _disposedValue = false;

    internal ExcelPivotCache(MsExcel.PivotCache pivotCache)
    {
        _pivotCache = pivotCache ?? throw new ArgumentNullException(nameof(pivotCache));
    }

    #region 基础属性
    public int Index => _pivotCache.Index;

    public object Parent => _pivotCache.Parent;

    public IExcelApplication Application => new ExcelApplication(_pivotCache.Application);

    public int SourceType => (int)_pivotCache.SourceType;

    public object SourceData => _pivotCache.SourceData;

    public int RecordCount => _pivotCache.RecordCount;

    public int Version => (int)_pivotCache.Version;
    #endregion

    #region 操作方法
    public void Refresh()
    {
        _pivotCache.Refresh();
    }

    public void Update()
    {
        Refresh();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_pivotCache != null)
                    Marshal.ReleaseComObject(_pivotCache);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _pivotCache = null;
        }
        _disposedValue = true;
    }

    ~ExcelPivotCache()
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
