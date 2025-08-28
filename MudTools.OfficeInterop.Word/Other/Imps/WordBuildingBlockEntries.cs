namespace MudTools.OfficeInterop.Word.Imps;

// <summary>
/// 对 <see cref="BuildingBlockEntries"/> COM 集合的封装实现类。
/// 提供类型安全的访问、遍历和资源管理。
/// </summary>
internal class WordBuildingBlockEntries : IWordBuildingBlockEntries
{
    private MsWord.BuildingBlockEntries _entries;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM 构建基块集合初始化封装实例。
    /// </summary>
    /// <param name="entries">原始的 BuildingBlockEntries COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 entries 为 null 时抛出。</exception>
    internal WordBuildingBlockEntries(MsWord.BuildingBlockEntries entries)
    {
        _entries = entries ?? throw new ArgumentNullException(nameof(entries));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public int Count
    {
        get
        {
            return _entries?.Count ?? 0;
        }
    }

    #endregion

    #region 索引器实现

    /// <inheritdoc />
    public IWordBuildingBlock? this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;
            var comBlock = _entries.Item(index);
            if (comBlock == null) return null;

            return new WordBuildingBlock(comBlock);
        }
    }
    #endregion

    #region 方法实现

    /// <inheritdoc />
    public IWordBuildingBlock Add(string name, WdBuildingBlockTypes type, string category, string value, WdDocPartInsertOptions insertOptions)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("名称不能为空。", nameof(name));
        if (string.IsNullOrWhiteSpace(category)) throw new ArgumentException("类别不能为空。", nameof(category));
        value ??= string.Empty;

        try
        {
            // 获取当前活动文档的内容范围，用于设置临时内容
            var app = _entries.Application;
            var doc = app.ActiveDocument;
            var range = doc.Content;
            range.Text = value;

            var newBlock = _entries.Add(
                Name: name,
                Type: (MsWord.WdBuildingBlockTypes)(int)type,
                Category: category,
                Range: range,
                InsertOptions: (MsWord.WdDocPartInsertOptions)(int)insertOptions
            );

            // 清空临时内容
            range.Text = string.Empty;

            return new WordBuildingBlock(newBlock);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加构建基块 '{name}'。", ex);
        }
    }

    /// <inheritdoc />
    public List<string> GetNames()
    {
        var names = new List<string>();

        if (_disposedValue || _entries == null) return names;

        for (int i = 1; i <= Count; i++)
        {
            var block = _entries.Item(i);
            if (block?.Name != null)
            {
                names.Add(block.Name.ToString());
            }
        }
        return names;
    }

    /// <inheritdoc />
    public void Clear()
    {
        if (_disposedValue || _entries == null) return;

        try
        {
            // 倒序删除，防止索引错乱
            for (int i = Count; i >= 1; i--)
            {
                var block = _entries.Item(i);
                block?.Delete();
            }
        }
        catch (COMException)
        {
            // 忽略整体异常
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源和托管资源。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 显式调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_entries != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_entries);
                }
                catch (InvalidComObjectException)
                {
                    // 已释放
                }
                finally
                {
                    _entries = null;
                }
            }
        }

        _disposedValue = true;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region 枚举器实现

    /// <inheritdoc />
    public IEnumerator<IWordBuildingBlock> GetEnumerator()
    {
        if (_disposedValue) yield break;

        for (int i = 1; i <= Count; i++)
        {
            var item = this[i];
            if (item != null)
            {
                yield return item;
            }
        }
    }

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}