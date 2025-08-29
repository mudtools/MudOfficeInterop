using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TableStyle 的实现类。
/// </summary>
internal class WordTableStyle : IWordTableStyle
{
    private MsWord.TableStyle _tableStyle;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="tableStyle">原始 COM TableStyle 对象。</param>
    internal WordTableStyle(MsWord.TableStyle tableStyle)
    {
        _tableStyle = tableStyle ?? throw new ArgumentNullException(nameof(tableStyle));
        _disposedValue = false;
    }

    #region 属性实现

    public WdTableDirection TableDirection
    {
        get => (WdTableDirection)(int)_tableStyle?.TableDirection;
        set
        {
            if (_tableStyle != null)
                _tableStyle.TableDirection = (MsWord.WdTableDirection)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordBorder? LeftBorder
    {
        get => _tableStyle?.Borders[MsWord.WdBorderType.wdBorderLeft] != null
            ? new WordBorder(_tableStyle.Borders[MsWord.WdBorderType.wdBorderLeft])
            : null;
    }

    /// <inheritdoc/>
    public IWordBorder? RightBorder
    {
        get => _tableStyle?.Borders[MsWord.WdBorderType.wdBorderRight] != null
            ? new WordBorder(_tableStyle.Borders[MsWord.WdBorderType.wdBorderRight])
            : null;
    }

    /// <inheritdoc/>
    public IWordBorder? TopBorder
    {
        get => _tableStyle?.Borders[MsWord.WdBorderType.wdBorderTop] != null
            ? new WordBorder(_tableStyle.Borders[MsWord.WdBorderType.wdBorderTop])
            : null;
    }

    /// <inheritdoc/>
    public IWordBorder? BottomBorder
    {
        get => _tableStyle?.Borders[MsWord.WdBorderType.wdBorderBottom] != null
            ? new WordBorder(_tableStyle.Borders[MsWord.WdBorderType.wdBorderBottom])
            : null;
    }

    /// <inheritdoc/>
    public IWordBorder? HorizontalBorder
    {
        get => _tableStyle?.Borders[MsWord.WdBorderType.wdBorderHorizontal] != null
            ? new WordBorder(_tableStyle.Borders[MsWord.WdBorderType.wdBorderHorizontal])
            : null;
    }

    /// <inheritdoc/>
    public IWordBorder? VerticalBorder
    {
        get => _tableStyle?.Borders[MsWord.WdBorderType.wdBorderVertical] != null
            ? new WordBorder(_tableStyle.Borders[MsWord.WdBorderType.wdBorderVertical])
            : null;
    }

    public bool AllowBreakAcrossPage
    {
        get => _tableStyle.AllowBreakAcrossPage == 1;
        set
        {
            if (_tableStyle != null)
                _tableStyle.AllowBreakAcrossPage = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool AllowPageBreaks
    {
        get => _tableStyle.AllowPageBreaks;
        set
        {
            if (_tableStyle != null)
                _tableStyle.AllowPageBreaks = value;
        }
    }

    /// <inheritdoc/>
    public IWordShading Shading => _tableStyle?.Shading != null ? new WordShading(_tableStyle.Shading) : null;

    /// <inheritdoc/>
    public WdRowAlignment Alignment
    {
        get => (WdRowAlignment)(int)_tableStyle?.Alignment;
        set
        {
            if (_tableStyle != null)
                _tableStyle.Alignment = (MsWord.WdRowAlignment)(int)value;
        }
    }

    public int ColumnStripe
    {
        get => _tableStyle?.ColumnStripe ?? 0;
        set
        {
            if (_tableStyle != null)
                _tableStyle.ColumnStripe = value;
        }
    }

    public int RowStripe
    {
        get => _tableStyle?.RowStripe ?? 0;
        set
        {
            if (_tableStyle != null)
                _tableStyle.RowStripe = value;
        }
    }

    /// <inheritdoc/>
    public float LeftIndent
    {
        get => _tableStyle?.LeftIndent ?? 0f;
        set
        {
            if (_tableStyle != null)
                _tableStyle.LeftIndent = value;
        }
    }

    /// <inheritdoc/>
    public float LeftPadding
    {
        get => _tableStyle?.LeftPadding ?? 0f;
        set
        {
            if (_tableStyle != null)
                _tableStyle.LeftPadding = value;
        }
    }

    /// <inheritdoc/>
    public float RightPadding
    {
        get => _tableStyle?.RightPadding ?? 0f;
        set
        {
            if (_tableStyle != null)
                _tableStyle.RightPadding = value;
        }
    }

    /// <inheritdoc/>
    public float TopPadding
    {
        get => _tableStyle?.TopPadding ?? 0f;
        set
        {
            if (_tableStyle != null)
                _tableStyle.TopPadding = value;
        }
    }

    /// <inheritdoc/>
    public float BottomPadding
    {
        get => _tableStyle?.BottomPadding ?? 0f;
        set
        {
            if (_tableStyle != null)
                _tableStyle.BottomPadding = value;
        }
    }

    public float Spacing
    {
        get => _tableStyle?.Spacing ?? 0f;
        set
        {
            if (_tableStyle != null)
                _tableStyle.Spacing = value;
        }
    }
    #endregion

    public IWordConditionalStyle Condition(WdConditionCode conditionCode)
    {
        if (_tableStyle == null)
            return null;

        var conditionalStyle = _tableStyle.Condition((MsWord.WdConditionCode)(int)conditionCode);
        return new WordConditionalStyle(conditionalStyle);
    }

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
            // 释放底纹对象
            if (_tableStyle?.Shading != null)
            {
                Marshal.ReleaseComObject(_tableStyle.Shading);
            }
            // 释放边框集合
            if (_tableStyle?.Borders != null)
            {
                Marshal.ReleaseComObject(_tableStyle.Borders);
            }
            // 释放表格样式对象本身
            if (_tableStyle != null)
            {
                Marshal.ReleaseComObject(_tableStyle);
                _tableStyle = null;
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