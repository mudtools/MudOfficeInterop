namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 <see cref="MsWord.BuildingBlockType"/> COM 对象的封装实现类。
/// 提供安全的属性访问、子集合封装和 COM 资源释放机制。
/// </summary>
internal class WordBuildingBlockType : IWordBuildingBlockType
{
    private MsWord.BuildingBlockType _type;
    private IWordCategories? _categories;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM BuildingBlockType 对象初始化封装实例。
    /// </summary>
    /// <param name="type">原始的 BuildingBlockType COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 type 为 null 时抛出。</exception>
    internal WordBuildingBlockType(MsWord.BuildingBlockType type)
    {
        _type = type ?? throw new ArgumentNullException(nameof(type));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public int Index
    {
        get
        {
            return _type?.Index ?? 0;
        }
    }

    /// <inheritdoc />
    public string Name
    {
        get
        {
            return _type?.Name ?? string.Empty;
        }
    }

    /// <inheritdoc />
    public IWordCategories Categories
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(WordBuildingBlockType));

            if (_categories == null)
            {
                var comCategories = (_type?.Categories) ?? throw new InvalidOperationException("无法获取该类型下的 Categories 集合。");
                _categories = new WordCategories(comCategories);
            }
            return _categories;
        }
    }

    /// <inheritdoc />
    public int TotalBlockCount
    {
        get
        {
            int totalCount = 0;

            // 遍历所有类别，累加 BuildingBlocks 数量
            foreach (var category in Categories)
            {
                using var blocks = category.BuildingBlocks;
                totalCount += blocks.Count;
            }

            return totalCount;
        }
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

        if (disposing)
        {
            // 释放子集合
            _categories?.Dispose();
            _categories = null;

            // 释放自身 COM 对象
            if (_type != null)
            {
                Marshal.ReleaseComObject(_type);
                _type = null;
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