//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 修订对象的封装实现类。
/// </summary>
internal class WordRevision : IWordRevision
{
    private MsWord.Revision _revision;
    private bool _disposedValue;

    internal WordRevision(MsWord.Revision revision)
    {
        _revision = revision ?? throw new ArgumentNullException(nameof(revision));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _revision != null ? new WordApplication(_revision.Application) : null;

    /// <inheritdoc/>
    public object Parent => _revision?.Parent;

    /// <inheritdoc/>
    public string Author => _revision?.Author ?? string.Empty;

    /// <inheritdoc/>
    public DateTime Date => _revision?.Date ?? DateTime.MinValue;

    /// <inheritdoc/>
    public WdRevisionType Type => _revision?.Type != null ? (WdRevisionType)(int)_revision?.Type : WdRevisionType.wdNoRevision;

    /// <inheritdoc/>
    public IWordRange Range => _revision?.Range != null ? new WordRange(_revision.Range) : null;

    /// <inheritdoc/>
    public string FormatDescription
    {
        get => _revision?.FormatDescription ?? string.Empty;
    }

    /// <inheritdoc/>
    public int Index => _revision?.Index ?? 0;

    /// <inheritdoc/>
    public IWordStyle Style => _revision?.Style != null ? new WordStyle(_revision.Style) : null;

    /// <inheritdoc/>
    public string RangeText => _revision?.Range?.Text ?? string.Empty;

    /// <inheritdoc/>
    public IWordRange MovedRange => _revision?.MovedRange != null ? new WordRange(_revision.MovedRange) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Accept()
    {
        _revision?.Accept();
    }

    /// <inheritdoc/>
    public void Reject()
    {
        _revision?.Reject();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _revision != null)
        {
            Marshal.ReleaseComObject(_revision);
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