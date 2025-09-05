namespace MudTools.OfficeInterop.Word.Imps;

// <summary>
/// 对 <see cref="AutoCorrectEntry"/> COM 对象的封装实现类。
/// 提供安全的属性访问和资源释放机制。
/// </summary>
internal class WordAutoCorrectEntry : IWordAutoCorrectEntry
{
    private MsWord.AutoCorrectEntry _entry;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM AutoCorrectEntry 对象初始化封装实例。
    /// </summary>
    /// <param name="entry">原始的 AutoCorrectEntry COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 entry 为 null 时抛出。</exception>
    internal WordAutoCorrectEntry(MsWord.AutoCorrectEntry entry)
    {
        _entry = entry ?? throw new ArgumentNullException(nameof(entry));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public string Name
    {
        get
        {
            return _entry?.Name ?? string.Empty;
        }
    }

    /// <inheritdoc />
    public string Value
    {
        get
        {
            return _entry?.Value ?? string.Empty;
        }
        set
        {
            value ??= string.Empty;
            _entry.Value = value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc />
    public void Delete()
    {
        if (_disposedValue || _entry == null) return;

        try
        {
            _entry.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("删除自动更正条目失败。", ex);
        }
    }

    public void Apply(IWordRange range)
    {
        if (range == null) return;
        try
        {
            _entry.Apply(((WordRange)range)._range);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("应用自动更正条目失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源（COM 对象）。
    /// </summary>
    /// <param name="disposing">为 true 表示由 Dispose() 显式调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _entry != null)
        {
            try
            {
                Marshal.ReleaseComObject(_entry);
            }
            catch (InvalidComObjectException)
            {
                // 忽略已释放对象
            }
            finally
            {
                _entry = null;
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
}