//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel SeriesCollection 对象的二次封装实现类
/// 实现 IExcelSeriesCollection 接口
/// </summary>
internal class ExcelSeriesCollection : IExcelSeriesCollection
{
    private MsExcel.SeriesCollection _seriesCollection;
    private bool _disposedValue = false;

    internal ExcelSeriesCollection(MsExcel.SeriesCollection seriesCollection)
    {
        _seriesCollection = seriesCollection ?? throw new ArgumentNullException(nameof(seriesCollection));
    }

    #region 基础属性
    public int Count => _seriesCollection.Count;

    public IExcelSeries this[int index]
    {
        get
        {
            MsExcel.Series series = _seriesCollection.Item(index);
            return new ExcelSeries(series);
        }
    }
    public object Parent => _seriesCollection.Parent;

    public IExcelApplication Application => new ExcelApplication(_seriesCollection.Application);
    #endregion

    #region 创建和添加
    public IExcelSeries Add()
    {
        MsExcel.Series newSeries = _seriesCollection.NewSeries();
        return new ExcelSeries(newSeries);
    }

    public IExcelSeries CreateSeries(IExcelRange source, int rowcol = 1, bool seriesLabels = false, bool categoryLabels = false)
    {
        ExcelRange comSource = source as ExcelRange;

        MsExcel.Series newSeries = _seriesCollection.Add(
            comSource.InternalRange,
            (MsExcel.XlRowCol)rowcol,
            seriesLabels,
            categoryLabels
        );
        return new ExcelSeries(newSeries);
    }
    #endregion

    #region 操作方法  

    public void Delete(int index)
    {
        try
        {
            MsExcel.Series seriesToDelete = _seriesCollection.Item(index);
            seriesToDelete?.Delete();
        }
        catch
        {

        }
    }

    public void Delete(IExcelSeries series)
    {
        if (series is ExcelSeries excelSeries && excelSeries != null)
        {
            excelSeries.InternalComObject.Delete();
        }
    }
    #endregion

    #region IEnumerable<IExcelSeries> Support
    public IEnumerator<IExcelSeries> GetEnumerator()
    {
        for (int i = 1; i <= _seriesCollection.Count; i++)
        {
            yield return new ExcelSeries(_seriesCollection.Item(i) as MsExcel.Series);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
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
                if (_seriesCollection != null)
                    Marshal.ReleaseComObject(_seriesCollection);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _seriesCollection = null;
        }
        _disposedValue = true;
    }

    ~ExcelSeriesCollection()
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
