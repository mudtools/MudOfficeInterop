//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Columns 的封装实现类。
/// </summary>
internal class WordColumns : IWordColumns
{
    private MsWord.Columns _columns;
    private bool _disposedValue;

    internal WordColumns(MsWord.Columns columns)
    {
        _columns = columns ?? throw new ArgumentNullException(nameof(columns));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _columns != null ? new WordApplication(_columns.Application) : null;

    /// <inheritdoc/>
    public object Parent => _columns?.Parent;

    /// <inheritdoc/>
    public int Count => _columns?.Count ?? 0;

    /// <inheritdoc/>
    public IWordColumn First => _columns?.Count > 0 ? new WordColumn(_columns[1]) : null;

    /// <inheritdoc/>
    public IWordColumn Last => _columns?.Count > 0 ? new WordColumn(_columns[_columns.Count]) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordColumn this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comColumn = _columns[index];
                return new WordColumn(comColumn);
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
    public IWordColumn Add(object beforeColumn = null)
    {
        if (_columns == null) return null;

        try
        {
            MsWord.Column newColumn;
            if (beforeColumn != null)
            {
                var targetColumn = (beforeColumn as WordColumn)?._column;
                if (targetColumn != null)
                {
                    newColumn = _columns.Add(targetColumn);
                }
                else
                {
                    newColumn = _columns.Add();
                }
            }
            else
            {
                newColumn = _columns.Add();
            }

            return new WordColumn(newColumn);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法添加列。", ex);
        }
    }

    /// <inheritdoc/>
    public void Delete(int index)
    {
        if (index < 1 || index > Count) return;

        _columns[index].Delete();
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
                _columns[i].Delete();
            }
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_columns == null) return;

        // 从后往前删除，避免索引变化
        for (int i = Count; i >= 1; i--)
        {
            _columns[i].Delete();
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
    public List<IWordColumn> Insert(int index, int count = 1)
    {
        var insertedColumns = new List<IWordColumn>();
        if (_columns == null || index < 1 || index > Count + 1 || count <= 0) return insertedColumns;

        try
        {
            for (int i = 0; i < count; i++)
            {
                MsWord.Column newColumn;
                if (index <= Count)
                {
                    var targetColumn = _columns[index];
                    newColumn = _columns.Add(targetColumn);
                }
                else
                {
                    newColumn = _columns.Add();
                }
                insertedColumns.Add(new WordColumn(newColumn));
            }
        }
        catch
        {
            // 插入失败返回已插入的列
        }

        return insertedColumns;
    }

    /// <inheritdoc/>
    public List<IWordColumn> GetRange(int startIndex, int endIndex)
    {
        var columns = new List<IWordColumn>();
        if (_columns == null) return columns;

        int validStart = Math.Max(1, Math.Min(startIndex, Count));
        int validEnd = Math.Max(validStart, Math.Min(endIndex, Count));

        try
        {
            for (int i = validStart; i <= validEnd; i++)
            {
                columns.Add(new WordColumn(_columns[i]));
            }
        }
        catch
        {
            // 获取范围失败返回已获取的列
        }

        return columns;
    }

    /// <inheritdoc/>
    public void Copy()
    {
        try
        {
            _columns.Application.Selection.Copy();
        }
        catch
        {
            // 复制失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Cut()
    {
        try
        {
            _columns.Application.Selection.Cut();
        }
        catch
        {
            // 剪切失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void Paste()
    {
        try
        {
            _columns.Application.Selection.Paste();
        }
        catch
        {
            // 粘贴失败忽略异常
        }
    }

    /// <inheritdoc/>
    public void AutoFit()
    {
        var table = _columns.Parent as MsWord.Table;
        table?.AutoFitBehavior(MsWord.WdAutoFitBehavior.wdAutoFitContent);
    }

    /// <inheritdoc/>
    public void SetBordersForAll(WdLineStyle lineStyle = WdLineStyle.wdLineStyleSingle,
                          WdLineWidth lineWidth = WdLineWidth.wdLineWidth100pt,
                          WdColor color = WdColor.wdColorAutomatic)
    {
        if (_columns == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var column = _columns[i];
                if (column?.Borders != null)
                {
                    column.Borders.Enable = 1;
                    foreach (MsWord.Border border in column.Borders)
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
        if (_columns == null) return;

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var column = _columns[i];
                if (column?.Shading != null)
                {
                    column.Shading.Texture = (MsWord.WdTextureIndex)(int)pattern;
                    if (foregroundColor != WdColor.wdColorAutomatic)
                        column.Shading.ForegroundPatternColor = (MsWord.WdColor)(int)foregroundColor;
                    if (backgroundColor != WdColor.wdColorWhite)
                        column.Shading.BackgroundPatternColor = (MsWord.WdColor)(int)backgroundColor;

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
    public IEnumerator<IWordColumn> GetEnumerator()
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

        if (disposing && _columns != null)
        {
            Marshal.ReleaseComObject(_columns);
            _columns = null;
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