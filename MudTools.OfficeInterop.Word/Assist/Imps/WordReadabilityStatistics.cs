//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 可读性统计信息集合的封装实现类。
/// </summary>
internal class WordReadabilityStatistics : IWordReadabilityStatistics
{
    private MsWord.ReadabilityStatistics _readabilityStatistics;
    private bool _disposedValue;

    internal WordReadabilityStatistics(MsWord.ReadabilityStatistics readabilityStatistics)
    {
        _readabilityStatistics = readabilityStatistics ?? throw new ArgumentNullException(nameof(readabilityStatistics));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _readabilityStatistics != null ? new WordApplication(_readabilityStatistics.Application) : null;

    /// <inheritdoc/>
    public object Parent => _readabilityStatistics?.Parent;

    /// <inheritdoc/>
    public int Count => _readabilityStatistics?.Count ?? 0;

    /// <inheritdoc/>
    public IWordReadabilityStatistic this[int index]
    {
        get
        {
            if (index < 1 || index > Count || _readabilityStatistics == null) return null;
            try
            {
                var comStat = _readabilityStatistics[index];
                return comStat != null ? new WordReadabilityStatistic(comStat) : null;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }

    /// <inheritdoc/>
    public IWordReadabilityStatistic this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name) || _readabilityStatistics == null) return null;
            try
            {
                var comStat = _readabilityStatistics[name];
                return comStat != null ? new WordReadabilityStatistic(comStat) : null;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _readabilityStatistics != null)
        {
            Marshal.ReleaseComObject(_readabilityStatistics);
            _readabilityStatistics = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable 实现

    public IEnumerator<IWordReadabilityStatistic> GetEnumerator()
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
}