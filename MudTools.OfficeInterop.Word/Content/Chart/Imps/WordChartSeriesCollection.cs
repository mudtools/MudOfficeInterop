//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.SeriesCollection 的封装实现类。
/// </summary>
internal class WordChartSeriesCollection : IWordChartSeriesCollection
{
    private MsWord.SeriesCollection _seriesCollection;
    private bool _disposedValue;

    internal WordChartSeriesCollection(MsWord.SeriesCollection seriesCollection)
    {
        _seriesCollection = seriesCollection ?? throw new ArgumentNullException(nameof(seriesCollection));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _seriesCollection != null ? new WordApplication(_seriesCollection.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _seriesCollection?.Parent;

    /// <inheritdoc/>
    public int Count => _seriesCollection?.Count ?? 0;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordChartSeries this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comSeries = _seriesCollection[index];
                return new WordChartSeries(comSeries);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IWordChartSeries this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            try
            {
                var comSeries = _seriesCollection[name];
                return comSeries != null ? new WordChartSeries(comSeries) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordChartSeries Add(object source, MsWord.XlRowCol rowcol, bool seriesLabels, bool categoryLabels, object bubbleSizes)
    {
        try
        {
            var newSeries = _seriesCollection.Add(source, rowcol, seriesLabels, categoryLabels, bubbleSizes);
            return new WordChartSeries(newSeries);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加数据系列。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordChartSeries NewSeries()
    {
        try
        {
            var newSeries = _seriesCollection.NewSeries();
            return new WordChartSeries(newSeries);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法创建新系列。", ex);
        }
    }

    /// <inheritdoc/>
    public bool Contains(string name)
    {
        if (_disposedValue || string.IsNullOrWhiteSpace(name)) return false;

        try
        {
            return _seriesCollection[name] != null;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public List<string> GetNames()
    {
        var names = new List<string>();
        for (int i = 1; i <= Count; i++)
        {
            var series = _seriesCollection[i];
            if (series?.Name != null)
                names.Add(series.Name.ToString());
        }
        return names;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordChartSeries> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _seriesCollection != null)
        {
            Marshal.ReleaseComObject(_seriesCollection);
            _seriesCollection = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}