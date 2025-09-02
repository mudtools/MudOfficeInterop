//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Cell 的封装实现类。
/// </summary>
internal class WordCell : IWordCell
{
    internal MsWord.Cell _cell;
    private bool _disposedValue;

    internal WordCell(MsWord.Cell cell)
    {
        _cell = cell ?? throw new ArgumentNullException(nameof(cell));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _cell != null ? new WordApplication(_cell.Application) : null;

    /// <inheritdoc/>
    public object Parent => _cell?.Parent;

    /// <inheritdoc/>
    public int RowIndex => _cell?.RowIndex ?? 0;

    /// <inheritdoc/>
    public int ColumnIndex => _cell?.ColumnIndex ?? 0;

    /// <inheritdoc/>
    public IWordRange? Range => _cell?.Range != null ? new WordRange(_cell.Range) : null;

    /// <inheritdoc/>
    public IWordRow? Row => _cell?.Row != null ? new WordRow(_cell.Row) : null;

    /// <inheritdoc/>
    public IWordColumn? Column => _cell?.Column != null ? new WordColumn(_cell.Column) : null;

    /// <inheritdoc/>
    public float Width
    {
        get => _cell?.Width ?? 0f;
        set
        {
            if (_cell != null)
                _cell.Width = value;
        }
    }

    /// <inheritdoc/>
    public float Height
    {
        get => _cell?.Height ?? 0f;
        set
        {
            if (_cell != null)
                _cell.Height = value;
        }
    }

    /// <inheritdoc/>
    public WdCellVerticalAlignment VerticalAlignment
    {
        get => _cell?.VerticalAlignment != null ? (WdCellVerticalAlignment)(int)_cell?.VerticalAlignment : WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        set
        {
            if (_cell != null) _cell.VerticalAlignment = (MsWord.WdCellVerticalAlignment)(int)value;
        }
    }

    /// <inheritdoc/>
    public IWordBorders Borders => _cell?.Borders != null ? new WordBorders(_cell.Borders) : null;

    /// <inheritdoc/>
    public IWordShading Shading => _cell?.Shading != null ? new WordShading(_cell.Shading) : null;

    /// <inheritdoc/>
    public float PreferredWidth
    {
        get => _cell?.PreferredWidth ?? 0f;
        set
        {
            if (_cell != null)
                _cell.PreferredWidth = value;
        }
    }

    /// <inheritdoc/>
    public WdPreferredWidthType PreferredWidthType
    {
        get => _cell?.PreferredWidthType != null ? (WdPreferredWidthType)(int)_cell?.PreferredWidthType : WdPreferredWidthType.wdPreferredWidthAuto;
        set
        {
            if (_cell != null) _cell.PreferredWidthType = (MsWord.WdPreferredWidthType)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool FitText
    {
        get => _cell?.FitText ?? false;
        set
        {
            if (_cell != null)
                _cell.FitText = value;
        }
    }

    /// <inheritdoc/>
    public float LeftPadding
    {
        get => _cell?.LeftPadding ?? 0f;
        set
        {
            if (_cell != null)
                _cell.LeftPadding = value;
        }
    }

    /// <inheritdoc/>
    public float RightPadding
    {
        get => _cell?.RightPadding ?? 0f;
        set
        {
            if (_cell != null)
                _cell.RightPadding = value;
        }
    }

    /// <inheritdoc/>
    public float TopPadding
    {
        get => _cell?.TopPadding ?? 0f;
        set
        {
            if (_cell != null)
                _cell.TopPadding = value;
        }
    }

    /// <inheritdoc/>
    public float BottomPadding
    {
        get => _cell?.BottomPadding ?? 0f;
        set
        {
            if (_cell != null)
                _cell.BottomPadding = value;
        }
    }

    /// <inheritdoc/>
    public IWordTables Tables => _cell?.Tables != null ? new WordTables(_cell.Tables) : null;

    /// <inheritdoc/>
    public IWordFont Font => _cell?.Range?.Font != null ? new WordFont(_cell.Range.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat => _cell?.Range?.ParagraphFormat != null ? new WordParagraphFormat(_cell.Range.ParagraphFormat) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _cell?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _cell?.Delete();
    }

    /// <inheritdoc/>
    public void Merge(IWordCell mergeTo)
    {
        if (_cell == null || mergeTo == null) return;

        try
        {
            var targetCell = (mergeTo as WordCell)?._cell;
            if (targetCell != null)
            {
                var cells = _cell.Application.Selection.Cells;
                cells.Merge();
            }
        }
        catch
        {
            // 合并失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Split(int numRows, int numColumns)
    {
        if (_cell == null || numRows <= 0 || numColumns <= 0) return;

        try
        {
            _cell.Split(numRows, numColumns);
        }
        catch
        {
            // 拆分失败忽略异常
        }
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return _cell?.Range?.Text?.TrimEnd('\r') ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_cell?.Range != null)
        {
            _cell.Range.Text = text ?? string.Empty;
        }
    }

    /// <inheritdoc/>
    public void ClearContents()
    {
        _cell?.Range?.Delete();
        if (_cell?.Range != null)
        {
            _cell.Range.Text = string.Empty;
        }
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _cell?.Range?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _cell?.Range?.Cut();
    }

    /// <inheritdoc/>
    public void Paste()
    {
        _cell?.Range?.Paste();
    }


    /// <inheritdoc/>
    public void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_cell?.Borders != null)
        {
            try
            {
                _cell.Borders.Enable = 1;
                foreach (MsWord.Border border in _cell.Borders)
                {
                    border.LineStyle = (MsWord.WdLineStyle)(int)lineStyle;
                    border.LineWidth = (MsWord.WdLineWidth)(int)lineWidth;
                    border.Color = (MsWord.WdColor)(int)color;
                }
            }
            catch
            {
                // 设置边框失败忽略异常
            }
        }
    }

    /// <inheritdoc/>
    public void RemoveBorders()
    {
        if (_cell == null)
            return;
        _cell.Borders.Enable = 0;
    }

    /// <inheritdoc/>
    public void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite)
    {
        if (_cell?.Shading != null)
        {
            try
            {
                _cell.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                if (foregroundColor != WdColor.wdColorAutomatic)
                    _cell.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                if (backgroundColor != WdColor.wdColorWhite)
                    _cell.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;
            }
            catch
            {
                // 设置底纹失败忽略异常
            }
        }
    }

    /// <inheritdoc/>
    public void RemoveShading()
    {
        if (_cell == null)
            return;
        _cell.Shading.Texture = MsWord.WdTextureIndex.wdTextureNone;
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_cell != null)
            {
                Marshal.ReleaseComObject(_cell);
            }
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
