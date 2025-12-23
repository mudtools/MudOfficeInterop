//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word.Row 的封装实现类。
/// </summary>
internal class WordRow : IWordRow
{
    internal MsWord.Row _row;

    internal MsWord.Row InternalComObject => _row;
    private bool _disposedValue;

    internal WordRow(MsWord.Row row)
    {
        _row = row ?? throw new ArgumentNullException(nameof(row));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _row != null ? new WordApplication(_row.Application) : null;

    /// <inheritdoc/>
    public object Parent => _row?.Parent;

    /// <inheritdoc/>
    public int Index => _row?.Index ?? 0;

    /// <inheritdoc/>
    public IWordRange Range => _row?.Range != null ? new WordRange(_row.Range) : null;

    /// <inheritdoc/>
    public IWordTable Table => _row?.Cells?.Parent as MsWord.Table != null ? new WordTable(_row.Cells.Parent as MsWord.Table) : null;

    /// <inheritdoc/>
    public float Height
    {
        get => _row?.Height ?? 0f;
        set
        {
            if (_row != null)
                _row.Height = value;
        }
    }

    /// <inheritdoc/>
    public WdRowHeightRule HeightRule
    {
        get => _row?.HeightRule.EnumConvert(WdRowHeightRule.wdRowHeightAuto) ?? WdRowHeightRule.wdRowHeightAuto;
        set
        {
            if (_row != null)
                _row.HeightRule = value.EnumConvert(MsWord.WdRowHeightRule.wdRowHeightAuto);
        }
    }

    /// <inheritdoc/>
    public bool AllowBreakAcrossPages
    {
        get => _row?.AllowBreakAcrossPages != null && _row?.AllowBreakAcrossPages == 1;
        set
        {
            if (_row != null)
                _row.AllowBreakAcrossPages = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public bool IsHeading
    {
        get => _row?.HeadingFormat != null && _row?.HeadingFormat == 1;
        set
        {
            if (_row != null)
                _row.HeadingFormat = value ? 1 : 0;
        }
    }

    /// <inheritdoc/>
    public IWordCells Cells => _row?.Cells != null ? new WordCells(_row.Cells) : null;

    /// <inheritdoc/>
    public IWordBorders Borders => _row?.Borders != null ? new WordBorders(_row.Borders) : null;

    /// <inheritdoc/>
    public IWordShading Shading => _row?.Shading != null ? new WordShading(_row.Shading) : null;


    /// <inheritdoc/>
    public float LeftIndent
    {
        get => _row?.LeftIndent ?? 0f;
        set
        {
            if (_row != null)
                _row.LeftIndent = value;
        }
    }

    /// <inheritdoc/>
    public IWordFont Font => _row?.Range?.Font != null ? new WordFont(_row.Range.Font) : null;

    /// <inheritdoc/>
    public IWordParagraphFormat ParagraphFormat => _row?.Range?.ParagraphFormat != null ? new WordParagraphFormat(_row.Range.ParagraphFormat) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _row?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _row?.Delete();
    }

    /// <inheritdoc/>
    public string GetText()
    {
        return _row?.Range?.Text?.TrimEnd('\r') ?? string.Empty;
    }

    /// <inheritdoc/>
    public void SetText(string text)
    {
        if (_row?.Range != null)
        {
            _row.Range.Text = text ?? string.Empty;
        }
    }

    /// <inheritdoc/>
    public void ClearContents()
    {
        _row?.Range?.Delete();
        if (_row?.Range != null)
        {
            _row.Range.Text = string.Empty;
        }
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _row?.Range?.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _row?.Range?.Cut();
    }

    /// <inheritdoc/>
    public void Paste()
    {
        _row?.Range?.Paste();
    }

    /// <inheritdoc/>
    public void Merge(IWordRow mergeTo)
    {
        if (_row == null || mergeTo == null) return;

        try
        {
            var targetRow = (mergeTo as WordRow)?._row;
            if (targetRow != null)
            {
                // 选择要合并的行范围
                var selection = _row.Application.Selection;
                selection.SetRange(_row.Range.Start, targetRow.Range.End);
                selection.Cells.Merge();
            }
        }
        catch
        {
            // 合并失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Split(int numRows)
    {
        if (_row == null || numRows <= 0) return;

        try
        {
            _row.Cells.Split(numRows);
        }
        catch
        {
            // 拆分失败忽略异常
        }
    }


    /// <inheritdoc/>
    public void SetBorders(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_row?.Borders != null)
        {
            try
            {
                _row.Borders.Enable = 1;
                foreach (MsWord.Border border in _row.Borders)
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
        if (_row == null)
            return;
        _row.Borders.Enable = 0;
    }

    /// <inheritdoc/>
    public void SetShading(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite)
    {
        if (_row?.Shading != null)
        {
            try
            {
                _row.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                if (foregroundColor != WdColor.wdColorAutomatic)
                    _row.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                if (backgroundColor != WdColor.wdColorWhite)
                    _row.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;
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
        if (_row == null)
            return;
        _row.Shading.Texture = MsWord.WdTextureIndex.wdTextureNone;
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_row != null)
            {
                Marshal.ReleaseComObject(_row);
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