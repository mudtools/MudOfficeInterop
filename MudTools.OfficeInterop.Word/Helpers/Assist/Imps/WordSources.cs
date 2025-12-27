//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示文档中所有书目源（引文）集合的封装实现类。
/// </summary>
internal class WordSources : IWordSources
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordSources));
    private MsWord.Sources _sources;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordSources"/> 类的新实例。
    /// </summary>
    /// <param name="sources">要封装的原始 COM Sources 对象。</param>
    internal WordSources(MsWord.Sources sources)
    {
        _sources = sources ?? throw new ArgumentNullException(nameof(sources));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _sources != null ? new WordApplication(_sources.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _sources?.Parent;

    /// <inheritdoc/>
    public int Creator => _sources?.Creator ?? 0;

    /// <inheritdoc/>
    public int Count => _sources?.Count ?? 0;

    #endregion

    #region 集合索引器实现 (Collection Indexer Implementation)

    /// <inheritdoc/>
    public IWordSource this[int index]
    {
        get
        {
            if (_sources == null) return null;
            try
            {
                var comSource = _sources[index];
                return comSource != null ? new WordSource(comSource) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 书目源集合方法实现 (Bibliography Sources Collection Methods Implementation)

    /// <inheritdoc/>
    public void Add(string data)
    {
        if (_sources == null || string.IsNullOrWhiteSpace(data)) return;
        try
        {
            _sources.Add(data);
        }
        catch (COMException ex)
        {
            log.Error($"Failed to add source: {ex.Message}");
            return;
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordSources"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _sources != null)
        {
            Marshal.ReleaseComObject(_sources);
            _sources = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordSources"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordSource> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordSource> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}