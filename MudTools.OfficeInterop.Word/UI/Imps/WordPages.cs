//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// <see cref="IWordPages"/> 接口的实现类，封装了 Microsoft.Office.Interop.Word.Panes 对象。
/// </summary>
internal class WordPages : IWordPages
{
    private MsWord.Pages _pages;
    private bool _disposedValue = false;

    /// <summary>
    /// 使用给定的 COM 对象初始化 <see cref="WordPages"/> 类的新实例。
    /// </summary>
    /// <param name="pages">原始的 Microsoft.Office.Interop.Word.Pages 对象。</param>
    /// <exception cref="ArgumentNullException">如果 <paramref name="pages"/> 为 null。</exception>
    internal WordPages(MsWord.Pages pages)
    {
        _pages = pages ?? throw new ArgumentNullException(nameof(pages));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _pages?.Application != null ? new WordApplication(_pages.Application) : null;

    /// <inheritdoc/>
    public int Count => _pages?.Count ?? 0;

    /// <inheritdoc/>
    public object Parent => _pages?.Parent;

    /// <inheritdoc/>
    public IWordPage this[int index]
    {
        get
        {
            if (_disposedValue || _pages == null)
            {
                return null; // 或抛出 ObjectDisposedException
            }
            if (index < 1 || index > Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");
            }
            try
            {
                var comPage = _pages[index];
                return new WordPage(comPage);
            }
            catch (COMException ex)
            {
                // 处理可能的 COM 错误，例如页面信息不可用
                System.Diagnostics.Debug.WriteLine($"Failed to retrieve page at index {index}: {ex.Message}");
                return null; // 或抛出自定义异常
            }
        }
    }

    #endregion    

    #region IEnumerable<IWordPage> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordPage> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++) // Word 集合索引通常从 1 开始
        {
            var page = this[i];
            if (page != null)
            {
                yield return page;
            }
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion // IEnumerable<IWordPage> 实现

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordPages"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放非托管资源 (COM 对象)
                if (_pages != null)
                {
                    Marshal.ReleaseComObject(_pages);
                    _pages = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放由 <see cref="WordPages"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion // IDisposable 实现
}