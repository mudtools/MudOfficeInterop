//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示文档中所有页眉和页脚集合的封装实现类。
/// </summary>
internal class WordHeadersFooters : IWordHeadersFooters
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordEditors));

    private MsWord.HeadersFooters _headersFooters;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordHeadersFooters"/> 类的新实例。
    /// </summary>
    /// <param name="headersFooters">要封装的原始 COM HeadersFooters 对象。</param>
    internal WordHeadersFooters(MsWord.HeadersFooters headersFooters)
    {
        _headersFooters = headersFooters ?? throw new ArgumentNullException(nameof(headersFooters));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _headersFooters != null ? new WordApplication(_headersFooters.Application) : null;

    /// <inheritdoc/>
    public object Parent => _headersFooters?.Parent;

    /// <inheritdoc/>
    public int Creator => _headersFooters?.Creator ?? 0;

    /// <inheritdoc/>
    public int Count => _headersFooters?.Count ?? 0;

    #endregion

    #region 集合索引器实现 (Collection Indexer Implementation)

    /// <inheritdoc/>
    public IWordHeaderFooter this[WdHeaderFooterIndex index]
    {
        get
        {
            if (_headersFooters == null) return null;
            try
            {
                var comHeaderFooter = _headersFooters[(MsWord.WdHeaderFooterIndex)(int)index];
                return comHeaderFooter != null ? new WordHeaderFooter(comHeaderFooter) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordHeadersFooters"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _headersFooters != null)
        {
            Marshal.ReleaseComObject(_headersFooters);
            _headersFooters = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordHeadersFooters"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordHeaderFooter> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordHeaderFooter> GetEnumerator()
    {
        foreach (var item in _headersFooters)
        {
            if (item != null && item is MsWord.HeaderFooter headerFooter)
            {
                yield return new WordHeaderFooter(headerFooter);
            }
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
