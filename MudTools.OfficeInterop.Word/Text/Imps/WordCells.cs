//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
// <summary>
/// Word.Cells 的封装实现类。
/// </summary>
internal class WordCells : IWordCells
{
    private MsWord.Cells _cells;
    private bool _disposedValue;

    internal WordCells(MsWord.Cells cells)
    {
        _cells = cells ?? throw new ArgumentNullException(nameof(cells));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _cells != null ? new WordApplication(_cells.Application) : null;

    /// <inheritdoc/>
    public object Parent => _cells?.Parent;

    /// <inheritdoc/>
    public int Count => _cells?.Count ?? 0;

    /// <inheritdoc/>
    public IWordCell First => _cells?.Count > 0 ? new WordCell(_cells[1]) : null;

    /// <inheritdoc/>
    public IWordCell Last => _cells?.Count > 0 ? new WordCell(_cells[_cells.Count]) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordCell this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comCell = _cells[index];
                return new WordCell(comCell);
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordCell Add(object beforeCell = null)
    {
        if (_cells == null) return null;

        try
        {
            MsWord.Cell newCell;
            if (beforeCell != null)
            {
                var targetCell = (beforeCell as WordCell)?._cell;
                if (targetCell != null)
                {
                    newCell = _cells.Add(targetCell);
                }
                else
                {
                    newCell = _cells.Add();
                }
            }
            else
            {
                newCell = _cells.Add();
            }

            return new WordCell(newCell);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加单元格。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        _cells[index].Delete();
    }

    /// <inheritdoc/>
    public void DeleteRange(int startIndex, int count)
    {
        if (startIndex < 1 || count <= 0) return;

        int endIndex = Math.Min(startIndex + count - 1, Count);
        if (startIndex <= endIndex)
        {
            // 从后往前删除，避免索引变化
            for (int i = endIndex; i >= startIndex; i--)
            {
                _cells[i].Delete();
            }
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_cells == null) return;

        // 从后往前删除，避免索引变化
        for (int i = Count; i >= 1; i--)
        {
            _cells[i].Delete();
        }
    }

    /// <inheritdoc/>
    public List<int> GetIndexes()
    {
        var indexes = new List<int>();
        for (int i = 1; i <= Count; i++)
        {
            indexes.Add(i);
        }
        return indexes;
    }

    /// <inheritdoc/>
    public void Merge(int startIndex, int endIndex)
    {
        if (_cells == null || startIndex < 1 || endIndex > Count || startIndex >= endIndex) return;

        var startCell = _cells[startIndex];
        var endCell = _cells[endIndex];

        // 选择要合并的单元格范围
        var range = startCell.Range.Duplicate;
        range.End = endCell.Range.End;
        range.Select();

        // 执行合并
        _cells.Application.Selection.Cells.Merge();
    }

    /// <inheritdoc/>
    public void Split(int numRows, int numColumns)
    {
        if (_cells == null || numRows <= 0 || numColumns <= 0) return;

        for (int i = 1; i <= Count; i++)
        {
            _cells[i].Split(numRows, numColumns);
        }
    }

    /// <inheritdoc/>
    public List<IWordCell> GetRange(int startIndex, int endIndex)
    {
        var cells = new List<IWordCell>();
        if (_cells == null) return cells;

        int validStart = Math.Max(1, Math.Min(startIndex, Count));
        int validEnd = Math.Max(validStart, Math.Min(endIndex, Count));

        for (int i = validStart; i <= validEnd; i++)
        {
            cells.Add(new WordCell(_cells[i]));
        }

        return cells;
    }

    /// <inheritdoc/>
    public void Select()
    {
        try
        {
            // 选择所有单元格
            if (Count > 0)
            {
                var firstCell = _cells[1];
                var lastCell = _cells[Count];
                var range = firstCell.Range.Duplicate;
                range.End = lastCell.Range.End;
                range.Select();
            }
        }
        catch
        {
            // 选择失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Copy()
    {
        _cells.Application.Selection.Copy();
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _cells.Application.Selection.Cut();
    }

    /// <inheritdoc/>
    public void Paste()
    {
        _cells.Application.Selection.Paste();
    }

    /// <inheritdoc/>
    public void ClearContents()
    {
        if (_cells == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                _cells[i].Range.Delete();
                _cells[i].Range.Text = string.Empty;
            }
        }
        catch
        {
            // 清除内容失败忽略异常
        }
    }


    /// <inheritdoc/>
    public void AutoFit()
    {
        var table = _cells.Parent as MsWord.Table;
        table?.AutoFitBehavior(MsWord.WdAutoFitBehavior.wdAutoFitContent);
    }

    /// <inheritdoc/>
    public void SetBordersForAll(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_cells == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var cell = _cells[i];
                if (cell?.Borders != null)
                {
                    cell.Borders.Enable = 1;
                    foreach (MsWord.Border border in cell.Borders)
                    {
                        border.LineStyle = (MsWord.WdLineStyle)(int)lineStyle;
                        border.LineWidth = (MsWord.WdLineWidth)(int)lineWidth;
                        border.Color = (MsWord.WdColor)(int)color;
                    }
                }
            }
        }
        catch
        {
            // 设置边框失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void SetShadingForAll(WdTextureIndex pattern = WdTextureIndex.wdTextureNone,
                          WdColor foregroundColor = WdColor.wdColorAutomatic,
                          WdColor backgroundColor = WdColor.wdColorWhite)
    {
        if (_cells == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var cell = _cells[i];
                if (cell?.Shading != null)
                {
                    cell.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                    if (foregroundColor != WdColor.wdColorAutomatic)
                        cell.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                    if (backgroundColor != WdColor.wdColorWhite)
                        cell.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;
                }
            }
        }
        catch
        {
            // 设置底纹失败忽略异常
        }
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordCell> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _cells != null)
        {
            Marshal.ReleaseComObject(_cells);
            _cells = null;
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