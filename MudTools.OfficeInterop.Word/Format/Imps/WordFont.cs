namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Font 的实现类。
/// </summary>
internal class WordFont : IWordFont
{
    private MsWord.Font _font;
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

    /// <inheritdoc/>
    public bool Underline
    {
        get => _font?.Underline != MsWord.WdUnderline.wdUnderlineNone;
        set
        {
            if (_font != null)
                _font.Underline = value
                    ? MsWord.WdUnderline.wdUnderlineSingle
                    : MsWord.WdUnderline.wdUnderlineNone;
        }
    }

    /// <inheritdoc/>
    public WdColor Color
    {
        get => (WdColor)(int)_font?.Color;
        set
        {
            if (_font != null)
                _font.Color = (MsWord.WdColor)(int)value;
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

    #endregion

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