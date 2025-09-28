
namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Replacement 的实现类。
/// </summary>
internal class WordReplacement : IWordReplacement
{
    private MsWord.Replacement? _replacement;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="replacement">原始 COM Replacement 对象。</param>
    internal WordReplacement(MsWord.Replacement replacement)
    {
        _replacement = replacement ?? throw new ArgumentNullException(nameof(replacement));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public string Text
    {
        get
        {
            return _replacement?.Text ?? string.Empty;
        }
        set
        {
            if (_replacement != null)
            {
                _replacement.Text = value ?? string.Empty;
            }
        }
    }

    /// <inheritdoc/>
    public IWordFont? Font =>
        _replacement?.Font != null ? new WordFont(_replacement.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat? ParagraphFormat =>
        _replacement?.ParagraphFormat != null ? new WordParagraphFormat(_replacement.ParagraphFormat) : null;

    public IWordFrame? Frame
    =>
    _replacement?.Frame != null ? new WordFrame(_replacement.Frame) : null;

    public WdLanguageID LanguageID
    {
        get => _replacement != null ? _replacement.LanguageID.EnumConvert(WdLanguageID.wdLanguageNone) : WdLanguageID.wdLanguageNone;
        set
        {
            if (_replacement != null)
                _replacement.LanguageID = value.EnumConvert(MsWord.WdLanguageID.wdLanguageNone);
        }
    }

    public WdBuiltinStyle Style
    {
        get => _replacement != null ? _replacement.get_Style().ObjectConvertEnum(WdBuiltinStyle.wdStyleNormal) : WdBuiltinStyle.wdStyleNormal;
        set
        {
            if (_replacement != null)
                _replacement.set_Style(value.EnumConvert(MsWord.WdBuiltinStyle.wdStyleNormal));
        }
    }

    public int NoProofing
    {
        get
        {
            return _replacement != null ? _replacement.NoProofing : 0;
        }
        set
        {
            if (_replacement != null)
                _replacement.NoProofing = value;
        }
    }

    /// <inheritdoc/>
    public float LineSpacing
    {
        get
        {

            return _replacement?.ParagraphFormat?.LineSpacing ?? 0f;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.LineSpacing = value;

            }
        }
    }

    /// <inheritdoc/>
    public WdLineSpacing LineSpacingRule
    {
        get
        {

            return _replacement != null ? _replacement.ParagraphFormat.LineSpacingRule.EnumConvert(WdLineSpacing.wdLineSpaceSingle) : WdLineSpacing.wdLineSpaceSingle;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.LineSpacingRule = value.EnumConvert(MsWord.WdLineSpacing.wdLineSpaceSingle);

            }
        }
    }

    /// <inheritdoc/>
    public float SpaceBefore
    {
        get
        {

            return _replacement?.ParagraphFormat?.SpaceBefore ?? 0f;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.SpaceBefore = value;

            }
        }
    }

    /// <inheritdoc/>
    public float SpaceAfter
    {
        get
        {

            return _replacement?.ParagraphFormat?.SpaceAfter ?? 0f;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.SpaceAfter = value;

            }
        }
    }

    /// <inheritdoc/>
    public float FirstLineIndent
    {
        get
        {

            return _replacement?.ParagraphFormat?.FirstLineIndent ?? 0f;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.FirstLineIndent = value;

            }
        }
    }

    /// <inheritdoc/>
    public float LeftIndent
    {
        get
        {

            return _replacement?.ParagraphFormat?.LeftIndent ?? 0f;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.LeftIndent = value;

            }
        }
    }

    /// <inheritdoc/>
    public float RightIndent
    {
        get
        {

            return _replacement?.ParagraphFormat?.RightIndent ?? 0f;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.RightIndent = value;

            }
        }
    }

    /// <inheritdoc/>
    public MsWord.WdParagraphAlignment Alignment
    {
        get
        {

            return _replacement?.ParagraphFormat?.Alignment ?? MsWord.WdParagraphAlignment.wdAlignParagraphLeft;
        }
        set
        {
            if (_replacement?.ParagraphFormat != null)
            {
                _replacement.ParagraphFormat.Alignment = value;

            }
        }
    }

    /// <inheritdoc/>
    public float CharacterSpacing
    {
        get
        {

            return _replacement?.Font?.Spacing ?? 0f;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Spacing = value;

            }
        }
    }

    /// <inheritdoc/>
    public int CharacterScaling
    {
        get
        {

            return _replacement?.Font?.Scaling ?? 100;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Scaling = value;

            }
        }
    }

    /// <inheritdoc/>
    public int Position
    {
        get
        {

            return _replacement?.Font?.Position ?? 0;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Position = value;

            }
        }
    }

    /// <inheritdoc/>
    public float FontSize
    {
        get
        {

            return _replacement?.Font?.Size ?? 0f;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Size = value;

            }
        }
    }

    /// <inheritdoc/>
    public string FontName
    {
        get
        {

            return _replacement?.Font?.Name ?? string.Empty;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Name = value;

            }
        }
    }

    /// <inheritdoc/>
    public bool Bold
    {
        get
        {

            return _replacement?.Font?.Bold == 1;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Bold = value ? 1 : 0;

            }
        }
    }

    /// <inheritdoc/>
    public bool Italic
    {
        get
        {

            return _replacement?.Font?.Italic == 1;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Italic = value ? 1 : 0;

            }
        }
    }

    /// <inheritdoc/>
    public bool Underline
    {
        get
        {

            return _replacement?.Font?.Underline != MsWord.WdUnderline.wdUnderlineNone;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Underline = value ?
                    MsWord.WdUnderline.wdUnderlineSingle :
                    MsWord.WdUnderline.wdUnderlineNone;

            }
        }
    }

    /// <inheritdoc/>
    public bool Superscript
    {
        get
        {

            return _replacement?.Font?.Superscript == 1;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Superscript = value ? 1 : 0;

            }
        }
    }

    /// <inheritdoc/>
    public bool Subscript
    {
        get
        {

            return _replacement?.Font?.Subscript == 1;
        }
        set
        {
            if (_replacement?.Font != null)
            {
                _replacement.Font.Subscript = value ? 1 : 0;

            }
        }
    }

    /// <inheritdoc/>
    public int Highlight
    {
        get => _replacement != null ? _replacement.Highlight : 0;


        set
        {
            if (_replacement != null)
            {
                _replacement.Highlight = value;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void ClearFormatting()
    {
        if (_replacement != null)
        {
            try
            {
                _replacement.ClearFormatting();

            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法清除替换格式。", ex);
            }
        }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放替换对象本身
            if (_replacement != null)
            {
                Marshal.ReleaseComObject(_replacement);
                _replacement = null;
            }
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