//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 超链接集合的封装实现类。
/// </summary>
internal class WordHyperlinks : IWordHyperlinks
{
    private MsWord.Hyperlinks _hyperlinks;
    private bool _disposedValue;

    internal WordHyperlinks(MsWord.Hyperlinks hyperlinks)
    {
        _hyperlinks = hyperlinks ?? throw new ArgumentNullException(nameof(hyperlinks));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _hyperlinks != null ? new WordApplication(_hyperlinks.Application) : null;

    /// <inheritdoc/>
    public int Count => _hyperlinks?.Count ?? 0;

    /// <inheritdoc/>
    public IWordHyperlink this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;
            var comHyperlink = _hyperlinks[index];
            return new WordHyperlink(comHyperlink);
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordHyperlink Add(IWordRange anchor, object address, object subAddress, object screenTip, object textToDisplay, object target)
    {
        if (anchor == null) throw new ArgumentNullException(nameof(anchor));

        try
        {
            var range = anchor as WordRange;
            var newHyperlink = _hyperlinks.Add(range.InternalComObject, address, subAddress, screenTip, textToDisplay, target);
            return new WordHyperlink(newHyperlink);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加超链接。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete()
    {
        if (_hyperlinks == null) return;

        // 从后往前删除，避免索引变化问题
        for (int i = Count; i >= 1; i--)
        {
            _hyperlinks[i]?.Delete();
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _hyperlinks != null)
        {
            Marshal.ReleaseComObject(_hyperlinks);
            _hyperlinks = null;
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

    public IEnumerator<IWordHyperlink> GetEnumerator()
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