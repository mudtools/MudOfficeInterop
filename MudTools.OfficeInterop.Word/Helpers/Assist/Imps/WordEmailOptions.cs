//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 表示 Microsoft Word 中用于电子邮件的全局首选项的封装实现类。
/// </summary>
internal class WordEmailOptions : IWordEmailOptions
{
    private MsWord.EmailOptions _emailOptions;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordEmailOptions"/> 类的新实例。
    /// </summary>
    /// <param name="emailOptions">要封装的原始 COM EmailOptions 对象。</param>
    internal WordEmailOptions(MsWord.EmailOptions emailOptions)
    {
        _emailOptions = emailOptions ?? throw new ArgumentNullException(nameof(emailOptions));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)

    /// <inheritdoc/>
    public IWordApplication? Application => _emailOptions != null ? new WordApplication(_emailOptions.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _emailOptions?.Parent;

    /// <inheritdoc/>
    public int Creator => _emailOptions?.Creator ?? 0;

    #endregion

    #region 电子邮件选项属性实现 (Email Options Properties Implementation)

    /// <inheritdoc/>
    public bool UseThemeStyle
    {
        get => _emailOptions?.UseThemeStyle ?? false;
        set { if (_emailOptions != null) _emailOptions.UseThemeStyle = value; }
    }

    /// <inheritdoc/>
    public string ThemeName
    {
        get => _emailOptions?.ThemeName ?? string.Empty;
        set { if (_emailOptions != null) _emailOptions.ThemeName = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public IWordStyle PlainTextStyle
    {
        get => _emailOptions?.PlainTextStyle != null ? new WordStyle(_emailOptions?.PlainTextStyle) : null;
    }


    /// <inheritdoc/>
    public bool RelyOnCSS
    {
        get => _emailOptions?.RelyOnCSS ?? false;
        set { if (_emailOptions != null) _emailOptions.RelyOnCSS = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeReplaceQuotes
    {
        get => _emailOptions?.AutoFormatAsYouTypeReplaceQuotes ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeReplaceQuotes = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyBorders
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyBorders ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyBorders = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyBulletedLists
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyBulletedLists ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyBulletedLists = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyNumberedLists
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyNumberedLists ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyNumberedLists = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyHeadings
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyHeadings ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyHeadings = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeFormatListItemBeginning
    {
        get => _emailOptions?.AutoFormatAsYouTypeFormatListItemBeginning ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeFormatListItemBeginning = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeDefineStyles
    {
        get => _emailOptions?.AutoFormatAsYouTypeDefineStyles ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeDefineStyles = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeReplaceSymbols
    {
        get => _emailOptions?.AutoFormatAsYouTypeReplaceSymbols ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeReplaceSymbols = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeReplaceOrdinals
    {
        get => _emailOptions?.AutoFormatAsYouTypeReplaceOrdinals ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeReplaceOrdinals = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeReplaceFractions
    {
        get => _emailOptions?.AutoFormatAsYouTypeReplaceFractions ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeReplaceFractions = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeReplacePlainTextEmphasis
    {
        get => _emailOptions?.AutoFormatAsYouTypeReplacePlainTextEmphasis ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeReplaceHyperlinks
    {
        get => _emailOptions?.AutoFormatAsYouTypeReplaceHyperlinks ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeReplaceHyperlinks = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyTables
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyTables ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyTables = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyFirstIndents
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyFirstIndents ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyFirstIndents = value; }
    }

    /// <inheritdoc/>
    public bool AutoFormatAsYouTypeApplyDates
    {
        get => _emailOptions?.AutoFormatAsYouTypeApplyDates ?? false;
        set { if (_emailOptions != null) _emailOptions.AutoFormatAsYouTypeApplyDates = value; }
    }


    /// <inheritdoc/>
    public bool MarkComments
    {
        get => _emailOptions?.MarkComments ?? false;
        set { if (_emailOptions != null) _emailOptions.MarkComments = value; }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordEmailOptions"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _emailOptions != null)
        {
            Marshal.ReleaseComObject(_emailOptions);
            _emailOptions = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 释放由 <see cref="WordEmailOptions"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}