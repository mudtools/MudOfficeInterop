//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// 表格对象的封装实现类。
/// </summary>
internal class WordTable : IWordTable
{
    private MsWord.Table _table;

    internal MsWord.Table InternalComObject => _table;
    private bool _disposedValue;

    internal WordTable(MsWord.Table table)
    {
        _table = table ?? throw new ArgumentNullException(nameof(table));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _table != null ? new WordApplication(_table.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _table?.Parent;

    /// <inheritdoc/>
    public IWordRange? Range => _table?.Range != null ? new WordRange(_table.Range) : null;

    /// <inheritdoc/>
    public bool Uniform => _table?.Uniform ?? false;

    /// <inheritdoc/>
    public IWordRows? Rows => _table?.Rows != null ? new WordRows(_table.Rows) : null;

    /// <inheritdoc/>
    public IWordColumns? Columns => _table?.Columns != null ? new WordColumns(_table.Columns) : null;

    /// <inheritdoc/>
    public IWordTables? Tables => _table?.Tables != null ? new WordTables(_table.Tables) : null;

    /// <inheritdoc/>
    public IWordShading? Shading => _table?.Shading != null ? new WordShading(_table.Shading) : null;

    /// <inheritdoc/>
    public IWordBorders? Borders => _table?.Borders != null ? new WordBorders(_table.Borders) : null;

    public object? Style
    {
        get => _table?.get_Style();
        set
        {
            if (_table != null)
                _table.set_Style(value);
        }
    }

    /// <inheritdoc/>
    public bool AllowPageBreaks
    {
        get => _table?.AllowPageBreaks ?? false;
        set
        {
            if (_table != null)
                _table.AllowPageBreaks = value;
        }
    }

    /// <inheritdoc/>
    public bool AllowAutoFit
    {
        get => _table?.AllowAutoFit ?? false;
        set
        {
            if (_table != null)
                _table.AllowAutoFit = value;
        }
    }

    public bool ApplyStyleHeadingRows
    {
        get => _table?.ApplyStyleHeadingRows ?? false;
        set
        {
            if (_table != null)
                _table.ApplyStyleHeadingRows = value;
        }
    }

    public bool ApplyStyleLastRow
    {
        get => _table?.ApplyStyleLastRow ?? false;
        set
        {
            if (_table != null)
                _table.ApplyStyleLastRow = value;
        }
    }

    public bool ApplyStyleFirstColumn
    {
        get => _table?.ApplyStyleFirstColumn ?? false;
        set
        {
            if (_table != null)
                _table.ApplyStyleFirstColumn = value;
        }
    }

    public bool ApplyStyleLastColumn
    {
        get => _table?.ApplyStyleLastColumn ?? false;
        set
        {
            if (_table != null)
                _table.ApplyStyleLastColumn = value;
        }
    }

    public bool ApplyStyleRowBands
    {
        get => _table?.ApplyStyleRowBands ?? false;
        set
        {
            if (_table != null)
                _table.ApplyStyleRowBands = value;
        }
    }

    public bool ApplyStyleColumnBands
    {
        get => _table?.ApplyStyleColumnBands ?? false;
        set
        {
            if (_table != null)
                _table.ApplyStyleColumnBands = value;
        }
    }

    /// <inheritdoc/>
    public object? TableStyle
    {
        get => _table?.get_Style();
        set
        {
            if (_table != null)
                _table.set_Style(value);
        }
    }

    /// <inheritdoc/>
    public string Title
    {
        get => _table?.Title ?? string.Empty;
        set
        {
            if (_table != null)
                _table.Title = value;
        }
    }

    /// <inheritdoc/>
    public string Descr
    {
        get => _table?.Descr ?? string.Empty;
        set
        {
            if (_table != null)
                _table.Descr = value;
        }
    }

    /// <inheritdoc/>
    public int NestingLevel => _table?.NestingLevel ?? 0;

    /// <inheritdoc/>
    public float PreferredWidth
    {
        get => _table?.PreferredWidth ?? 0f;
        set
        {
            if (_table != null)
                _table.PreferredWidth = value;
        }
    }

    /// <inheritdoc/>
    public WdPreferredWidthType PreferredWidthType
    {
        get => _table != null ? _table.PreferredWidthType.EnumConvert(WdPreferredWidthType.wdPreferredWidthAuto) : WdPreferredWidthType.wdPreferredWidthAuto;
        set
        {
            if (_table != null) _table.PreferredWidthType = value.EnumConvert(MsWord.WdPreferredWidthType.wdPreferredWidthAuto);
        }
    }

    /// <inheritdoc/>
    public WdTableDirection TableDirection
    {
        get => _table?.TableDirection != null ? _table.TableDirection.EnumConvert(WdTableDirection.wdTableDirectionRtl) : WdTableDirection.wdTableDirectionRtl;
        set
        {
            if (_table != null) _table.TableDirection = value.EnumConvert(MsWord.WdTableDirection.wdTableDirectionRtl);
        }
    }

    /// <inheritdoc/>
    public float TopPadding
    {
        get => _table?.TopPadding ?? 0f;
        set
        {
            if (_table != null)
                _table.TopPadding = value;
        }
    }

    /// <inheritdoc/>
    public float BottomPadding
    {
        get => _table?.BottomPadding ?? 0f;
        set
        {
            if (_table != null)
                _table.BottomPadding = value;
        }
    }

    /// <inheritdoc/>
    public float LeftPadding
    {
        get => _table?.LeftPadding ?? 0f;
        set
        {
            if (_table != null)
                _table.LeftPadding = value;
        }
    }

    /// <inheritdoc/>
    public float RightPadding
    {
        get => _table?.RightPadding ?? 0f;
        set
        {
            if (_table != null)
                _table.RightPadding = value;
        }
    }

    /// <inheritdoc/>
    public float Spacing
    {
        get => _table?.Spacing ?? 0f;
        set
        {
            if (_table != null)
                _table.Spacing = value;
        }
    }

    /// <inheritdoc/>
    public string ID
    {
        get => _table?.ID ?? string.Empty;
        set
        {
            if (_table != null)
                _table.ID = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordCell? Cell(int rowIndex, int columnIndex)
    {
        if (_table == null) return null;
        try
        {
            var comCell = _table.Cell(rowIndex, columnIndex);
            return comCell != null ? new WordCell(comCell) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void AutoFitBehavior(MsWord.WdAutoFitBehavior behavior)
    {
        _table?.AutoFitBehavior(behavior);
    }

    /// <inheritdoc/>
    public IWordRange? ConvertToText(object separator)
    {
        if (_table == null) return null;
        try
        {
            var range = _table.ConvertToText(ref separator);
            return range != null ? new WordRange(range) : null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _table?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _table?.Select();
    }

    /// <inheritdoc/>
    public void Split(object beforeRow)
    {
        _table?.Split(ref beforeRow);
    }

    /// <inheritdoc/>
    public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object ascending)
    {
        _table?.Sort(ref excludeHeader, ref fieldNumber, ref sortFieldType, ref ascending);
    }

    /// <inheritdoc/>
    public void ApplyStyleDirectFormatting(string styleName)
    {
        if (_table == null || string.IsNullOrWhiteSpace(styleName)) return;
        _table.ApplyStyleDirectFormatting(styleName);
    }

    public void AutoFitBehavior(WdAutoFitBehavior behavior)
    {
        if (_table == null) return;
        _table.AutoFitBehavior(behavior.EnumConvert(MsWord.WdAutoFitBehavior.wdAutoFitContent));
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _table != null)
        {
            Marshal.ReleaseComObject(_table);
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