//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word 文档书签集合实现类
/// </summary>
internal class WordBookmarks : IWordBookmarks
{
    private readonly MsWord.Bookmarks _bookmarks;
    private bool _disposedValue;

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    public IWordApplication? Application => _bookmarks != null ? new WordApplication(_bookmarks.Application) : null;

    /// <summary>
    /// 获取父对象
    /// </summary>
    public object Parent => _bookmarks?.Parent;

    /// <summary>
    /// 获取书签数量
    /// </summary>
    public int Count => _bookmarks.Count;

    /// <summary>
    /// 构造函数
    /// </summary>
    /// <param name="bookmarks">COM Bookmarks 对象</param>
    internal WordBookmarks(MsWord.Bookmarks bookmarks)
    {
        _bookmarks = bookmarks ?? throw new ArgumentNullException(nameof(bookmarks));
        _disposedValue = false;
    }

    /// <summary>
    /// 根据索引获取书签
    /// </summary>
    /// <param name="index">书签索引</param>
    /// <returns>书签对象</returns>
    public IWordBookmark this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Index must be between 1 and {Count}.");

            try
            {
                var bookmark = _bookmarks[index];
                return new WordBookmark(bookmark);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get bookmark at index {index}.", ex);
            }
        }
    }

    /// <summary>
    /// 根据名称获取书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>书签对象</returns>
    public IWordBookmark this[string name]
    {
        get
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

            try
            {
                var bookmark = _bookmarks[name];
                return new WordBookmark(bookmark);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get bookmark with name '{name}'.", ex);
            }
        }
    }

    /// <summary>
    /// 添加书签
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <param name="range">书签范围</param>
    /// <returns>书签对象</returns>
    public IWordBookmark Add(string name, IWordRange range)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        try
        {
            // 注意：这里需要将 IWordRange 转换为 COM Range 对象
            // 由于缺少具体实现，这里使用占位符
            var comRange = GetComRange(range);
            var bookmark = _bookmarks.Add(name, comRange);
            return new WordBookmark(bookmark);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to add bookmark '{name}'.", ex);
        }
    }

    /// <summary>
    /// 删除书签
    /// </summary>
    /// <param name="name">书签名称</param>
    public void Delete(string name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name cannot be null or empty.", nameof(name));

        try
        {
            if (Exists(name))
            {
                _bookmarks[name].Delete();
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to delete bookmark '{name}'.", ex);
        }
    }

    /// <summary>
    /// 检查书签是否存在
    /// </summary>
    /// <param name="name">书签名称</param>
    /// <returns>是否存在</returns>
    public bool Exists(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;

        try
        {
            return _bookmarks.Exists(name);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>书签枚举器</returns>
    public IEnumerator<IWordBookmark> GetEnumerator()
    {
        try
        {
            var bookmarks = new List<IWordBookmark>();
            for (int i = 1; i <= Count; i++)
            {
                try
                {
                    bookmarks.Add(this[i]);
                }
                catch
                {
                    // 忽略获取失败的书签
                    continue;
                }
            }
            return bookmarks.GetEnumerator();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to enumerate bookmarks.", ex);
        }
    }

    /// <summary>
    /// 获取枚举器
    /// </summary>
    /// <returns>枚举器</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// 将 IWordRange 转换为 COM Range 对象
    /// </summary>
    /// <param name="range">IWordRange 对象</param>
    /// <returns>COM Range 对象</returns>
    private MsWord.Range GetComRange(IWordRange range)
    {
        // 这里需要具体的实现来获取 COM Range 对象
        // 由于缺少具体实现，返回 null 作为占位符
        return null;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing">是否正在 disposing</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;
        _disposedValue = true;
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}

