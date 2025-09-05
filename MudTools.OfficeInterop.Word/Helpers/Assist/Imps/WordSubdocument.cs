//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 子文档对象的封装实现类。
/// </summary>
internal class WordSubdocument : IWordSubdocument
{
    private MsWord.Subdocument _subdocument;
    private bool _disposedValue;

    internal WordSubdocument(MsWord.Subdocument subdocument)
    {
        _subdocument = subdocument ?? throw new ArgumentNullException(nameof(subdocument));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _subdocument != null ? new WordApplication(_subdocument.Application) : null;

    /// <inheritdoc/>
    public object Parent => _subdocument?.Parent;

    /// <inheritdoc/>
    public string Name => _subdocument?.Name ?? string.Empty;

    /// <inheritdoc/>
    public string Path => _subdocument?.Path ?? string.Empty;

    /// <inheritdoc/>
    public IWordRange Range => _subdocument?.Range != null ? new WordRange(_subdocument.Range) : null;

    /// <inheritdoc/>
    public bool Locked
    {
        get => _subdocument?.Locked ?? false;
        set
        {
            if (_subdocument != null)
                _subdocument.Locked = value;
        }
    }

    /// <inheritdoc/>
    public bool HasFile => _subdocument?.HasFile ?? false;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordDocument Open()
    {
        if (_subdocument == null) return null;

        try
        {
            var document = _subdocument.Open();
            return document != null ? new WordDocument(document) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _subdocument?.Delete();
    }

    /// <inheritdoc/>
    public void Split(IWordRange range)
    {

        _subdocument?.Split(((WordRange)range)._range);
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _subdocument != null)
        {
            Marshal.ReleaseComObject(_subdocument);
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