namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// 对 <see cref="MsWord.AutoCorrectEntries"/> COM 集合的封装实现类。
/// 提供对自动更正条目的安全访问、添加、查询和资源管理。
/// </summary>
internal class WordAutoCorrectEntries : IWordAutoCorrectEntries
{
    private MsWord.AutoCorrectEntries _entries;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM AutoCorrectEntries 集合初始化封装实例。
    /// </summary>
    /// <param name="entries">原始的 AutoCorrectEntries COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 entries 为 null 时抛出。</exception>
    internal WordAutoCorrectEntries(MsWord.AutoCorrectEntries entries)
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
    public IWordAutoCorrectEntry? this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;
            var comEntry = _entries[index];
            if (comEntry == null) return null;

            return new WordAutoCorrectEntry(comEntry);
        }
    }

    /// <inheritdoc />
    public IWordAutoCorrectEntry? this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            var comEntry = _entries[name];
            if (comEntry == null) return null;

            return new WordAutoCorrectEntry(comEntry);
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc />
    public IWordAutoCorrectEntry Add(string name, string value)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("名称不能为空。", nameof(name));
        value ??= string.Empty;

        try
        {
            var newEntry = _entries.Add(Name: name, Value: value);
            return new WordAutoCorrectEntry(newEntry);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加自动更正条目 '{name}'。", ex);
        }
    }

    /// <inheritdoc />
    public bool Contains(string name)
    {
        return _entries[name] != null;
    }

    /// <inheritdoc />
    public List<string> GetNames()
    {
        var names = new List<string>();

        if (_disposedValue || _entries == null) return names;

        for (int i = 1; i <= Count; i++)
        {
            var entry = _entries[i];
            if (entry?.Name != null)
            {
                names.Add(entry.Name.ToString());
            }
        }

        return names;
    }

    /// <inheritdoc />
    public void Clear()
    {
        if (_disposedValue || _entries == null) return;

        for (int i = Count; i >= 1; i--)
        {
            var entry = _entries[i];
            entry?.Delete();
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 显式调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _entries != null)
        {
            Marshal.ReleaseComObject(_entries);
            _entries = null;
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
    public IEnumerator<IWordAutoCorrectEntry> GetEnumerator()
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