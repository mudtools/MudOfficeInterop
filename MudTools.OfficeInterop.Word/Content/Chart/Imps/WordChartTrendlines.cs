//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Trendlines 的封装实现类。
/// </summary>
internal class WordChartTrendlines : IWordChartTrendlines
{
    private MsWord.Trendlines _trendlines;
    private bool _disposedValue;

    internal WordChartTrendlines(MsWord.Trendlines trendlines)
    {
        _trendlines = trendlines ?? throw new ArgumentNullException(nameof(trendlines));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _trendlines != null ? new WordApplication(_trendlines.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _trendlines?.Parent;

    /// <inheritdoc/>
    public int Count => _trendlines?.Count ?? 0;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordChartTrendline this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comTrendline = _trendlines[index];
                return new WordChartTrendline(comTrendline);
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
    public IWordChartTrendline Add(MsWord.XlTrendlineType type, object order, object period)
    {
        try
        {
            var newTrendline = _trendlines.Add(type, order, period);
            return new WordChartTrendline(newTrendline);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加趋势线。", ex);
        }
    }

    /// <inheritdoc/>
    public List<int> GetIndexes()
    {
        var indexes = new List<int>();
        for (int i = 1; i <= Count; i++)
        {
            indexes.Add(i);
        }
        return indexes;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordChartTrendline> GetEnumerator()
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

        if (disposing && _trendlines != null)
        {
            Marshal.ReleaseComObject(_trendlines);
            _trendlines = null;
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