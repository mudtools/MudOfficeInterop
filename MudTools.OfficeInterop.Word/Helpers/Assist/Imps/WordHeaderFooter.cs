//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示单个页眉或页脚的封装实现类。
/// </summary>
internal class WordHeaderFooter : IWordHeaderFooter
{
    private MsWord.HeaderFooter _headerFooter;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordHeaderFooter"/> 类的新实例。
    /// </summary>
    /// <param name="headerFooter">要封装的原始 COM HeaderFooter 对象。</param>
    internal WordHeaderFooter(MsWord.HeaderFooter headerFooter)
    {
        _headerFooter = headerFooter ?? throw new ArgumentNullException(nameof(headerFooter));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication Application => _headerFooter != null ? new WordApplication(_headerFooter.Application) : null;

    /// <inheritdoc/>
    public object Parent => _headerFooter?.Parent;

    /// <inheritdoc/>
    public int Creator => _headerFooter?.Creator ?? 0;

    #endregion

    #region 页眉/页脚属性实现 (Header/Footer Properties Implementation)

    /// <inheritdoc/>
    public IWordRange? Range =>
        _headerFooter?.Range != null ? new WordRange(_headerFooter.Range) : null;


    /// <inheritdoc/>
    public IWordPageNumbers? PageNumbers =>
        _headerFooter?.PageNumbers != null ? new WordPageNumbers(_headerFooter.PageNumbers) : null;

    /// <inheritdoc/>
    public WdHeaderFooterIndex Index =>
        _headerFooter?.Index != null ? (WdHeaderFooterIndex)(int)_headerFooter.Index : WdHeaderFooterIndex.wdHeaderFooterPrimary;

    /// <inheritdoc/>
    public IWordShapes? Shapes =>
        _headerFooter?.Shapes != null ? new WordShapes(_headerFooter.Shapes) : null;
    #endregion


    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordHeaderFooter"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _headerFooter != null)
        {
            Marshal.ReleaseComObject(_headerFooter);
            _headerFooter = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordHeaderFooter"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}