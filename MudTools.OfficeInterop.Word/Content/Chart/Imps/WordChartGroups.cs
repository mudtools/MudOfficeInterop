//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ChartGroups 的封装实现类。
/// </summary>
internal class WordChartGroups : IWordChartGroups
{
    private MsWord.ChartGroups _chartGroups;
    private bool _disposedValue;

    internal WordChartGroups(MsWord.ChartGroups chartGroups)
    {
        _chartGroups = chartGroups ?? throw new ArgumentNullException(nameof(chartGroups));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartGroups != null ? new WordApplication(_chartGroups.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartGroups?.Parent;

    /// <inheritdoc/>
    public int Count => _chartGroups?.Count ?? 0;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordChartGroup this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            var comChartGroup = _chartGroups[index];
            return new WordChartGroup(comChartGroup);
        }
    }

    #endregion

    #region 方法实现
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
    public IEnumerator<IWordChartGroup> GetEnumerator()
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

        if (disposing && _chartGroups != null)
        {
            Marshal.ReleaseComObject(_chartGroups);
            _chartGroups = null;
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