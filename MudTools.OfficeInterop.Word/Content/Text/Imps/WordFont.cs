//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Font 的实现类。
/// </summary>
internal class WordFont : IWordFont
{
    private MsWord.Font? _font;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="font">原始 COM Font 对象。</param>
    internal WordFont(MsWord.Font font)
    {
        _font = font ?? throw new ArgumentNullException(nameof(font));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _font != null ? new WordApplication(_font.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _font?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _font?.Name ?? string.Empty;
        set
        {
            if (_font != null)
                _font.Name = value;
        }
    }

    /// <inheritdoc/>
    public float Size
    {
        get => _font?.Size ?? 0f;
        set
        {
            if (_font != null)
                _font.Size = value;
        }
    }

    /// <inheritdoc/>
    public bool Bold
    {
        get => _font?.Bold == 1;
        set
        {
            if (_font != null)
                _font.Bold = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool Italic
    {
        get => _font?.Italic == 1;
        set
        {
            if (_font != null)
                _font.Italic = value ? 1 : 0;
        }
    }

    public bool Hidden
    {
        get => _font?.Hidden == 1;
        set
        {
            if (_font != null)
                _font.Hidden = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public WdUnderline Underline
    {
        get => _font != null ? _font.Underline.EnumConvert(WdUnderline.wdUnderlineNone) : WdUnderline.wdUnderlineNone;
        set
        {
            if (_font != null)
                _font.Underline = value.EnumConvert(MsWord.WdUnderline.wdUnderlineNone);
        }
    }

    /// <inheritdoc/>
    public WdColor Color
    {
        get => _font != null ? _font.Color.EnumConvert(WdColor.wdColorAutomatic) : WdColor.wdColorAutomatic;
        set
        {
            if (_font != null)
                _font.Color = value.EnumConvert(MsWord.WdColor.wdColorAutomatic);
        }
    }

    /// <inheritdoc/>
    public bool Superscript
    {
        get => _font?.Superscript == 1;
        set
        {
            if (_font != null)
                _font.Superscript = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool Subscript
    {
        get => _font?.Subscript == 1;
        set
        {
            if (_font != null)
                _font.Subscript = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public float Spacing
    {
        get => _font?.Spacing ?? 0f;
        set
        {
            if (_font != null)
                _font.Spacing = value;
        }
    }

    /// <inheritdoc/>
    public int Scaling
    {
        get => _font?.Scaling ?? 100;
        set
        {
            if (_font != null)
                _font.Scaling = value;
        }
    }

    /// <inheritdoc/>
    public int Position
    {
        get => _font?.Position ?? 0;
        set
        {
            if (_font != null)
                _font.Position = value;
        }
    }

    public string NameFarEast
    {
        get => _font?.NameFarEast ?? string.Empty;
        set
        {
            if (_font != null)
                _font.NameFarEast = value;
        }
    }

    public bool Outline
    {
        get => _font?.Outline == 1;
        set
        {
            if (_font != null)
                _font.Outline = value ? 1 : 0;
        }
    }

    public bool Shadow
    {
        get => _font?.Shadow == 1;
        set
        {
            if (_font != null)
                _font.Shadow = value ? 1 : 0;
        }
    }

    public bool DisableCharacterSpaceGrid
    {
        get => _font != null ? _font.DisableCharacterSpaceGrid : false;
        set
        {
            if (_font != null)
                _font.DisableCharacterSpaceGrid = value;
        }
    }

    public WdColorIndex ColorIndex
    {
        get => _font != null ? _font.ColorIndex.EnumConvert(WdColorIndex.wdAuto) : WdColorIndex.wdAuto;
        set
        {
            if (_font != null)
                _font.ColorIndex = value.EnumConvert(MsWord.WdColorIndex.wdAuto);
        }
    }

    public WdColor UnderlineColor
    {
        get => _font != null ? _font.UnderlineColor.EnumConvert(WdColor.wdColorAutomatic) : WdColor.wdColorAutomatic;
        set
        {
            if (_font != null)
                _font.UnderlineColor = value.EnumConvert(MsWord.WdColor.wdColorAutomatic);
        }
    }

    public WdNumberSpacing NumberSpacing
    {
        get => _font != null ? _font.NumberSpacing.EnumConvert(WdNumberSpacing.wdNumberSpacingDefault) : WdNumberSpacing.wdNumberSpacingDefault;
        set
        {
            if (_font != null)
                _font.NumberSpacing = value.EnumConvert(MsWord.WdNumberSpacing.wdNumberSpacingDefault);
        }
    }

    public WdStylisticSet StylisticSet
    {
        get => _font != null ? _font.StylisticSet.EnumConvert(WdStylisticSet.wdStylisticSetDefault) : WdStylisticSet.wdStylisticSetDefault;
        set
        {
            if (_font != null)
                _font.StylisticSet = value.EnumConvert(MsWord.WdStylisticSet.wdStylisticSetDefault);
        }
    }

    public IWordFont? Duplicate
    {
        get
        {
            if (_font != null)
                return new WordFont(_font.Duplicate);
            return null;
        }
    }

    public IWordBorders? Borders
    {
        get
        {
            if (_font != null)
                return new WordBorders(_font.Borders);
            return null;
        }
    }

    public IWordFillFormat? Fill
    {
        get
        {
            if (_font != null)
                return new WordFillFormat(_font.Fill);
            return null;
        }
    }

    public IWordGlowFormat? Glow
    {
        get
        {
            if (_font != null)
                return new WordGlowFormat(_font.Glow);
            return null;
        }
    }

    public IWordLineFormat? Line
    {
        get
        {
            if (_font != null)
                return new WordLineFormat(_font.Line);
            return null;
        }
    }

    public IWordReflectionFormat? Reflection
    {
        get
        {
            if (_font != null)
                return new WordReflectionFormat(_font.Reflection);
            return null;
        }
    }

    public IWordShading? Shading
    {
        get
        {
            if (_font != null)
                return new WordShading(_font.Shading);
            return null;
        }
    }

    public IWordColorFormat? TextColor
    {
        get
        {
            if (_font != null)
                return new WordColorFormat(_font.TextColor);
            return null;
        }
    }

    public IWordShadowFormat? TextShadow
    {
        get
        {
            if (_font != null)
                return new WordShadowFormat(_font.TextShadow);
            return null;
        }
    }

    public IWordThreeDFormat? ThreeD
    {
        get
        {
            if (_font != null)
                return new WordThreeDFormat(_font.ThreeD);
            return null;
        }
    }

    #endregion

    public void Grow()
    {
        if (_font != null)
            _font.Grow();
    }

    public void Reset()
    {
        if (_font != null)
            _font.Reset();
    }

    public void SetAsTemplateDefault()
    {
        if (_font != null)
            _font.SetAsTemplateDefault();
    }

    public void Shrink()
    {
        if (_font != null)
            _font.Shrink();
    }

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _font != null)
        {
            Marshal.ReleaseComObject(_font);
            _font = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}