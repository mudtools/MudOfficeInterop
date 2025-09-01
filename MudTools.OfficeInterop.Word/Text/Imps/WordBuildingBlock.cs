
namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 <see cref="MsWord.BuildingBlock"/> COM 对象的封装实现类。
/// 提供安全的属性访问和资源释放机制。
/// </summary>
internal class WordBuildingBlock : IWordBuildingBlock
{
    private MsWord.BuildingBlock _block;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM 构建基块对象初始化封装实例。
    /// </summary>
    /// <param name="block">原始的 BuildingBlock COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 block 为 null 时抛出。</exception>
    internal WordBuildingBlock(MsWord.BuildingBlock block)
    {
        _block = block ?? throw new ArgumentNullException(nameof(block));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public string Name => _block?.Name;

    /// <inheritdoc />
    public string Value
    {
        get
        {
            return _block?.Value;
        }
        set
        {
            _block.Value = value;
        }
    }

    /// <inheritdoc />
    public IWordCategory? Category => _block != null ? new WordCategory(_block.Category) : null;

    /// <inheritdoc />
    public string Type
    {
        get
        {
            try
            {
                return _block?.Type.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc />
    public void Delete()
    {
        if (_disposedValue || _block == null) return;

        try
        {
            _block.Delete();
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("删除构建基块失败。", ex);
        }
        finally
        {
            Dispose(); // 删除后立即释放引用
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放非托管资源和托管资源（如果需要）。
    /// </summary>
    /// <param name="disposing">为 true 时表示由 Dispose() 显式调用；false 表示由终结器调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _block != null)
        {
            try
            {
                Marshal.ReleaseComObject(_block);
            }
            catch (InvalidComObjectException)
            {
                // 忽略已释放的对象
            }
            finally
            {
                _block = null;
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