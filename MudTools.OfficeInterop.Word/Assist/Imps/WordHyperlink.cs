//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Core;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 超链接对象的封装实现类。
/// </summary>
internal class WordHyperlink : IWordHyperlink
{
    private MsWord.Hyperlink _hyperlink;
    private bool _disposedValue;

    internal WordHyperlink(MsWord.Hyperlink hyperlink)
    {
        _hyperlink = hyperlink ?? throw new ArgumentNullException(nameof(hyperlink));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _hyperlink != null ? new WordApplication(_hyperlink.Application) : null;

    /// <inheritdoc/>
    public object Parent => _hyperlink?.Parent;

    /// <inheritdoc/>
    public string TextToDisplay
    {
        get => _hyperlink?.TextToDisplay ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.TextToDisplay = value;
        }
    }

    /// <inheritdoc/>
    public string Address
    {
        get => _hyperlink?.Address ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.Address = value;
        }
    }

    /// <inheritdoc/>
    public string SubAddress
    {
        get => _hyperlink?.SubAddress ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.SubAddress = value;
        }
    }

    /// <inheritdoc/>
    public string ScreenTip
    {
        get => _hyperlink?.ScreenTip ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.ScreenTip = value;
        }
    }

    /// <inheritdoc/>
    public IWordRange Range => _hyperlink?.Range != null ? new WordRange(_hyperlink.Range) : null;

    /// <inheritdoc/>
    public MsoHyperlinkType Type => _hyperlink?.Type != null ? (MsoHyperlinkType)(int)_hyperlink?.Type : MsoHyperlinkType.msoHyperlinkRange;

    /// <inheritdoc/>
    public string Target
    {
        get => _hyperlink?.Target ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.Target = value;
        }
    }

    /// <inheritdoc/>
    public string EmailSubject
    {
        get => _hyperlink?.EmailSubject ?? string.Empty;
        set
        {
            if (_hyperlink != null)
                _hyperlink.EmailSubject = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Delete()
    {
        _hyperlink?.Delete();
    }

    /// <inheritdoc/>
    public void Follow()
    {
        _hyperlink?.Follow();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _hyperlink != null)
        {
            Marshal.ReleaseComObject(_hyperlink);
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