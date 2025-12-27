//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 ListGallery 对象集合的封装实现类。
/// </summary>
internal class WordListGalleries : IWordListGalleries
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordListGalleries));
    private MsWord.ListGalleries _listGalleries;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordListGalleries"/> 类的新实例。
    /// </summary>
    /// <param name="listGalleries">要封装的原始 COM ListGalleries 对象。</param>
    internal WordListGalleries(MsWord.ListGalleries listGalleries)
    {
        _listGalleries = listGalleries ?? throw new ArgumentNullException(nameof(listGalleries));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _listGalleries != null ? new WordApplication(_listGalleries.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _listGalleries?.Parent;

    /// <inheritdoc/>
    public int Count => _listGalleries?.Count ?? 0;

    /// <inheritdoc/>
    public IWordListGallery this[WdListGalleryType index]
    {
        get
        {
            if (_listGalleries == null) return null;
            try
            {
                var comListGallery = _listGalleries[(MsWord.WdListGalleryType)(int)index];
                return comListGallery != null ? new WordListGallery(comListGallery) : null;
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
    /// 释放由 <see cref="WordListGalleries"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _listGalleries != null)
        {
            Marshal.ReleaseComObject(_listGalleries);
            _listGalleries = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordListGalleries"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordListGallery> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordListGallery> GetEnumerator()
    {
        foreach (var item in _listGalleries)
        {
            if (item != null && item is ListGallery gallery)
            {
                yield return new WordListGallery(gallery);
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