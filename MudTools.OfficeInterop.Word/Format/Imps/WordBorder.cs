namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Border 的实现类。
/// </summary>
internal class WordBorder : IWordBorder
{
    private MsWord.Border _border;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="border">原始 COM Border 对象。</param>
    internal WordBorder(MsWord.Border border)
    {
        _border = border ?? throw new ArgumentNullException(nameof(border));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public bool Visible
    {
        get => _border != null && _border.Visible;
        set
        {
            if (_border != null)
                _border.Visible = value;
        }
    }

    /// <inheritdoc/>
    public WdLineStyle LineStyle
    {
        get => (WdLineStyle)(int)_border!.LineStyle;
        set
        {
            if (_border != null)
                _border.LineStyle = (MsWord.WdLineStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdLineWidth LineWidth
    {
        get => (WdLineWidth)(int)_border?.LineWidth;
        set
        {
            if (_border != null)
                _border.LineWidth = (MsWord.WdLineWidth)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColor Color
    {
        get => (WdColor)(int)_border?.Color;
        set
        {
            if (_border != null)
                _border.Color = (MsWord.WdColor)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColorIndex ColorIndex
    {
        get => (WdColorIndex)(int)_border?.ColorIndex;
        set
        {
            if (_border != null)
                _border.ColorIndex = (MsWord.WdColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdPageBorderArt ArtStyle
    {
        get => (WdPageBorderArt)(int)_border?.ArtStyle;
        set
        {
            if (_border != null)
                _border.ArtStyle = (MsWord.WdPageBorderArt)(int)value;
        }
    }

    /// <inheritdoc/>
    public int ArtWidth
    {
        get => _border?.ArtWidth ?? 0;
        set
        {
            if (_border != null)
                _border.ArtWidth = value;
        }
    }

    /// <inheritdoc/>
    public bool Inside => _border != null && _border.Inside;

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _border != null)
        {
            Marshal.ReleaseComObject(_border);
            _border = null;
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