namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// 对 <see cref="MsWord.Categories"/> COM 集合的封装实现类。
/// 表示某一构建基块类型（如页眉）下的所有类别集合。
/// 提供安全的访问、查询和资源管理机制。
/// </summary>
internal class WordCategories : IWordCategories
{
    private MsWord.Categories _categories;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM Categories 集合初始化封装实例。
    /// </summary>
    /// <param name="categories">原始的 Categories COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 categories 为 null 时抛出。</exception>
    internal WordCategories(MsWord.Categories categories)
    {
        _categories = categories ?? throw new ArgumentNullException(nameof(categories));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public int Count
    {
        get
        {
            return _categories?.Count ?? 0;
        }
    }

    #endregion

    #region 索引器实现

    /// <inheritdoc />
    public IWordCategory? this[int index]
    {
        get
        {
            if (index < 1 || index > Count)
                return null;

            var comCategory = _categories.Item(index);
            if (comCategory == null)
                return null;

            return new WordCategory(comCategory);
        }
    }

    /// <inheritdoc />
    public IWordCategory? this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            var comCategory = _categories.Item(name);
            if (comCategory == null)
                return null;

            return new WordCategory(comCategory);
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc />
    public bool Contains(string name)
    {
        if (_disposedValue || string.IsNullOrWhiteSpace(name))
            return false;

        return _categories.Item(name) != null;
    }

    /// <inheritdoc />
    public List<string> GetNames()
    {
        var names = new List<string>();

        if (_disposedValue || _categories == null)
            return names;

        for (int i = 1; i <= Count; i++)
        {
            var category = _categories.Item(i);
            if (category?.Name != null)
            {
                names.Add(category.Name.ToString());
            }
        }

        return names;
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源（COM 对象）和托管资源（如果需要）。
    /// </summary>
    /// <param name="disposing">为 true 表示由 Dispose() 显式调用；false 表示由终结器调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _categories != null)
        {
            Marshal.ReleaseComObject(_categories);
            _categories = null;
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
    public IEnumerator<IWordCategory> GetEnumerator()
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