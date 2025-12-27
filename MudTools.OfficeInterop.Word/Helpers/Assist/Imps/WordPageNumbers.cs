//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Word.Imps;
// <summary>
/// 表示文档中所有页码集合的封装实现类。
/// </summary>
internal class WordPageNumbers : IWordPageNumbers
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordPageNumbers));
    private MsWord.PageNumbers _pageNumbers;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordPageNumbers"/> 类的新实例。
    /// </summary>
    /// <param name="pageNumbers">要封装的原始 COM PageNumbers 对象。</param>
    internal WordPageNumbers(MsWord.PageNumbers pageNumbers)
    {
        _pageNumbers = pageNumbers ?? throw new ArgumentNullException(nameof(pageNumbers));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _pageNumbers != null ? new WordApplication(_pageNumbers.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _pageNumbers?.Parent;

    /// <inheritdoc/>
    public int Creator => _pageNumbers?.Creator ?? 0;

    /// <inheritdoc/>
    public int Count => _pageNumbers?.Count ?? 0;

    #endregion

    #region 集合索引器实现 (Collection Indexer Implementation)

    /// <inheritdoc/>
    public IWordPageNumber this[int index]
    {
        get
        {
            if (_pageNumbers == null || index < 1 || index > Count) return null;
            try
            {
                var comPageNumber = _pageNumbers[index];
                return comPageNumber != null ? new WordPageNumber(comPageNumber) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 页码集合属性实现 (Page Numbers Collection Properties Implementation)

    /// <inheritdoc/>
    public bool ShowFirstPageNumber
    {
        get => _pageNumbers?.ShowFirstPageNumber ?? false;
        set { if (_pageNumbers != null) _pageNumbers.ShowFirstPageNumber = value; }
    }

    /// <inheritdoc/>
    public int StartingNumber
    {
        get => _pageNumbers?.StartingNumber ?? 0;
        set { if (_pageNumbers != null) _pageNumbers.StartingNumber = value; }
    }

    #endregion

    #region 页码集合方法实现 (Page Numbers Collection Methods Implementation)

    /// <inheritdoc/>
    public IWordPageNumber Add(WdPageNumberAlignment alignment, int pageNumbers)
    {
        if (_pageNumbers == null) return null;
        try
        {
            var newPageNumber = _pageNumbers.Add((MsWord.WdPageNumberAlignment)alignment, pageNumbers);
            return newPageNumber != null ? new WordPageNumber(newPageNumber) : null;
        }
        catch (COMException ex)
        {
            log.Error($"Failed to add page number: {ex.Message}");
            return null;
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordPageNumbers"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _pageNumbers != null)
        {
            Marshal.ReleaseComObject(_pageNumbers);
            _pageNumbers = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordPageNumbers"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordPageNumber> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordPageNumber> GetEnumerator()
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