using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// <see cref="IWordAutoTextEntries"/> 接口的具体实现类，封装 Word 中的 AutoTextEntries 集合。
/// 支持枚举、添加、查找和删除自动图文集条目。
/// </summary>
internal class WordAutoTextEntries : IWordAutoTextEntries
{
    private AutoTextEntries _entries;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数：包装一个 AutoTextEntries 集合
    /// </summary>
    /// <param name="entries">Interop 中的 AutoTextEntries 集合</param>
    internal WordAutoTextEntries(AutoTextEntries entries)
    {
        _entries = entries ?? throw new ArgumentNullException(nameof(entries));
        _disposedValue = false;
    }

    #region 属性实现

    /// <summary>
    /// 获取集合中自动图文集条目的总数
    /// </summary>
    public int Count
    {
        get
        {
            return _entries?.Count ?? 0;
        }
    }

    #endregion

    #region 索引器实现

    /// <summary>
    /// 根据 1-based 索引获取封装后的自动图文集条目
    /// </summary>
    /// <param name="index">索引（从 1 开始）</param>
    /// <returns>封装对象；索引越界返回 null</returns>
    public IWordAutoTextEntry? this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comEntry = _entries[index];
                var wrapper = new WordAutoTextEntry(comEntry);
                return wrapper;
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// 根据名称获取封装后的自动图文集条目
    /// </summary>
    /// <param name="name">条目名称</param>
    /// <returns>封装对象；未找到返回 null</returns>
    public IWordAutoTextEntry? this[string name]
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(WordAutoTextEntries));
            if (string.IsNullOrWhiteSpace(name)) return null;

            try
            {
                var comEntry = _entries[name];
                if (comEntry == null) return null;

                var wrapper = new WordAutoTextEntry(comEntry);
                return wrapper;
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <summary>
    /// 向模板中添加一个新的自动图文集条目
    /// </summary>
    /// <param name="name">条目名称</param>
    /// <param name="value">插入的内容（文本）</param>
    /// <returns>新创建的封装条目对象</returns>
    public IWordAutoTextEntry Add(string name, string value)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("名称不能为空。", nameof(name));
        value ??= string.Empty;

        try
        {
            var range = _entries.Application.ActiveDocument.Content;
            range.Text = value;

            var newEntry = _entries.Add(name, range);
            range.Text = string.Empty;

            var wrapper = new WordAutoTextEntry(newEntry);
            return wrapper;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加自动图文集条目 '{name}'。", ex);
        }
    }

    /// <summary>
    /// 检查是否存在指定名称的自动图文集条目
    /// </summary>
    /// <param name="name">要查找的名称</param>
    /// <returns>是否存在</returns>
    public bool Contains(string name)
    {
        if (_disposedValue || string.IsNullOrWhiteSpace(name)) return false;
        try
        {
            return _entries[name] != null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取所有自动图文集条目的名称列表
    /// </summary>
    /// <returns>名称列表</returns>
    public List<string> GetNames()
    {
        var names = new List<string>();

        try
        {
            for (int i = 1; i <= Count; i++)
            {
                var entry = _entries[i];
                if (entry?.Name != null)
                    names.Add(entry.Name.ToString());
            }
        }
        catch
        {
            // 忽略遍历异常
        }

        return names;
    }

    /// <summary>
    /// 删除所有自动图文集条目（慎用！）
    /// </summary>
    public void Clear()
    {
        if (_disposedValue || _entries == null) return;

        try
        {
            for (int i = Count; i >= 1; i--)
            {
                try
                {
                    _entries[i]?.Delete();
                }
                catch { }
            }
        }
        catch (COMException)
        {
            // 忽略
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放集合本身
            if (_entries != null)
            {
                Marshal.ReleaseComObject(_entries);
                _entries = null;
            }
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    public IEnumerator<IWordAutoTextEntry> GetEnumerator()
    {
        for (int i = 0; i < Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}