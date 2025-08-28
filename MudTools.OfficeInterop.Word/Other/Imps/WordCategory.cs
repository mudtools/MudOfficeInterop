
using Microsoft.Office.Interop.Word;

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 <see cref="Category"/> COM 对象的封装实现类。
/// 提供安全的属性访问和 COM 资源释放机制。
/// </summary>
internal class WordCategory : IWordCategory
{
    private MsWord.Category _category;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM Category 对象初始化封装实例。
    /// </summary>
    /// <param name="category">原始的 Category COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 category 为 null 时抛出。</exception>
    internal WordCategory(MsWord.Category category)
    {
        _category = category ?? throw new ArgumentNullException(nameof(category));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public string Name
    {
        get => _category?.Name ?? string.Empty;
    }

    public IWordBuildingBlocks BuildingBlocks => _category != null ? new WordBuildingBlocks(_category.BuildingBlocks) : null;

    /// <inheritdoc />
    public int Index
    {
        get => _category.Index;
    }

    public IWordBuildingBlockType Type => _category != null ? new WordBuildingBlockType(_category.Type) : null;

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源（COM 对象）和托管资源（如果需要）。
    /// </summary>
    /// <param name="disposing">为 true 表示由 Dispose() 显式调用；false 表示由终结器调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _category != null)
        {
            Marshal.ReleaseComObject(_category);
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