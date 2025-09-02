//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Rows 的封装实现类。
/// </summary>
internal class WordRows : IWordRows
{
    private MsWord.Rows _rows;
    private bool _disposedValue;

    internal WordRows(MsWord.Rows rows)
    {
        _rows = rows ?? throw new ArgumentNullException(nameof(rows));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _rows != null ? new WordApplication(_rows.Application) : null;

    /// <inheritdoc/>
    public object Parent => _rows?.Parent;

    /// <inheritdoc/>
    public int Count => _rows?.Count ?? 0;

    /// <inheritdoc/>
    public IWordRow First => _rows?.Count > 0 ? new WordRow(_rows[1]) : null;

    /// <inheritdoc/>
    public IWordRow Last => _rows?.Count > 0 ? new WordRow(_rows[_rows.Count]) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordRow this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comRow = _rows[index];
                return new WordRow(comRow);
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
    public IWordRow Add(object beforeRow = null)
    {
        if (_rows == null) return null;

        try
        {
            MsWord.Row newRow;
            if (beforeRow != null)
            {
                var targetRow = (beforeRow as WordRow)?._row;
                if (targetRow != null)
                {
                    newRow = _rows.Add(targetRow);
                }
                else
                {
                    newRow = _rows.Add();
                }
            }
            else
            {
                newRow = _rows.Add();
            }

            return new WordRow(newRow);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加行。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        _rows[index].Delete();
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
                _rows[i].Delete();
            }
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_rows == null) return;

        // 从后往前删除，避免索引变化
        for (int i = Count; i >= 1; i--)
        {
            _rows[i].Delete();
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
    public List<IWordRow> Insert(int index, int count = 1)
    {
        var insertedRows = new List<IWordRow>();
        if (_rows == null || index < 1 || index > Count + 1 || count <= 0) return insertedRows;

        try
        {
            for (int i = 0; i < count; i++)
            {
                MsWord.Row newRow;
                if (index <= Count)
                {
                    var targetRow = _rows[index];
                    newRow = _rows.Add(targetRow);
                }
                else
                {
                    newRow = _rows.Add();
                }
                insertedRows.Add(new WordRow(newRow));
            }
        }
        catch
        {
            // 插入失败返回已插入的行
        }

        return insertedRows;
    }

    /// <inheritdoc/>
    public List<IWordRow> GetRange(int startIndex, int endIndex)
    {
        var rows = new List<IWordRow>();
        if (_rows == null) return rows;

        int validStart = Math.Max(1, Math.Min(startIndex, Count));
        int validEnd = Math.Max(validStart, Math.Min(endIndex, Count));

        try
        {
            for (int i = validStart; i <= validEnd; i++)
            {
                rows.Add(new WordRow(_rows[i]));
            }
        }
        catch
        {
            // 获取范围失败返回已获取的行
        }

        return rows;
    }

    /// <inheritdoc/>
    public void Select()
    {
        try
        {
            // 选择所有行
            if (Count > 0)
            {
                var firstRow = _rows[1];
                var lastRow = _rows[Count];
                var range = firstRow.Range.Duplicate;
                range.End = lastRow.Range.End;
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
        try
        {
            _rows.Application.Selection.Copy();
        }
        catch
        {
            // 复制失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Cut()
    {
        _rows.Application.Selection.Cut();
    }

    /// <inheritdoc/>
    public void Paste()
    {
        _rows.Application.Selection.Paste();
    }

    /// <inheritdoc/>
    public void ClearContents()
    {
        if (_rows == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                _rows[i].Range.Delete();
                _rows[i].Range.Text = string.Empty;
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
        try
        {
            var table = _rows.Parent as MsWord.Table;
            table?.AutoFitBehavior(MsWord.WdAutoFitBehavior.wdAutoFitContent);
        }
        catch
        {
            // 自动调整失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Merge(int startIndex, int endIndex)
    {
        if (_rows == null || startIndex < 1 || endIndex > Count || startIndex >= endIndex) return;

        try
        {
            var startRow = _rows[startIndex];
            var endRow = _rows[endIndex];

            // 选择要合并的行范围
            var range = startRow.Range.Duplicate;
            range.End = endRow.Range.End;
            range.Select();

            // 执行合并
            _rows.Application.Selection.Cells.Merge();
        }
        catch
        {
            // 合并失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Split(int numRows)
    {
        if (_rows == null || numRows <= 0) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                _rows[i].Cells.Split(numRows);
            }
        }
        catch
        {
            // 拆分失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void SetBordersForAll(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_rows == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var row = _rows[i];
                if (row?.Borders != null)
                {
                    row.Borders.Enable = 1;
                    foreach (MsWord.Border border in row.Borders)
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
        if (_rows == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var row = _rows[i];
                if (row?.Shading != null)
                {
                    row.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                    if (foregroundColor != WdColor.wdColorAutomatic)
                        row.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                    if (backgroundColor != WdColor.wdColorWhite)
                        row.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;

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
    public IEnumerator<IWordRow> GetEnumerator()
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

        if (disposing && _rows != null)
        {
            Marshal.ReleaseComObject(_rows);
            _rows = null;
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