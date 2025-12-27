//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表示单个页码的封装实现类。
/// </summary>
internal class WordPageNumber : IWordPageNumber
{
    private MsWord.PageNumber _pageNumber;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordPageNumber"/> 类的新实例。
    /// </summary>
    /// <param name="pageNumber">要封装的原始 COM PageNumber 对象。</param>
    internal WordPageNumber(MsWord.PageNumber pageNumber)
    {
        _pageNumber = pageNumber ?? throw new ArgumentNullException(nameof(pageNumber));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication? Application => _pageNumber != null ? new WordApplication(_pageNumber.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _pageNumber?.Parent;

    /// <inheritdoc/>
    public int Creator => _pageNumber?.Creator ?? 0;

    /// <inheritdoc/>
    public int Index => _pageNumber?.Index ?? 0;

    #endregion

    #region 页码属性实现 (Page Number Properties Implementation)

    /// <inheritdoc/>
    public WdPageNumberAlignment Alignment
    {
        get => _pageNumber?.Alignment != null ? (WdPageNumberAlignment)(int)_pageNumber.Alignment : WdPageNumberAlignment.wdAlignPageNumberCenter;
        set
        {
            if (_pageNumber != null) _pageNumber.Alignment = (MsWord.WdPageNumberAlignment)(int)value;
        }
    }

    #endregion

    #region 页码方法实现 (Page Number Methods Implementation)

    /// <inheritdoc/>
    public void Delete()
    {
        _pageNumber?.Delete();
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _pageNumber?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _pageNumber?.Cut();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _pageNumber?.Select();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordPageNumber"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _pageNumber != null)
        {
            Marshal.ReleaseComObject(_pageNumber);
            _pageNumber = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordPageNumber"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}