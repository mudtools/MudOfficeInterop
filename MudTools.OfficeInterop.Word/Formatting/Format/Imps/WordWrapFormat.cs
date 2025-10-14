
namespace MudTools.OfficeInterop.Word.Imps;


/// <summary>
/// IWordWrapFormat 接口的具体实现类。
/// </summary>
internal class WordWrapFormat : IWordWrapFormat
{
    private MsWord.WrapFormat _wrapFormat;
    private bool _disposedValue = false;

    /// <summary>
    /// 构造函数，使用原始的 COM 对象进行初始化。
    /// </summary>
    /// <param name="wrapFormat">原始的 Microsoft.Office.Interop.Word.WrapFormat 对象。</param>
    internal WordWrapFormat(MsWord.WrapFormat wrapFormat)
    {
        _wrapFormat = wrapFormat ?? throw new ArgumentNullException(nameof(wrapFormat));
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _wrapFormat?.Application != null ? new WordApplication(_wrapFormat.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _wrapFormat?.Parent;

    /// <inheritdoc/>
    public WdWrapType Type
    {
        get => _wrapFormat?.Type.EnumConvert(WdWrapType.wdWrapInline) ?? WdWrapType.wdWrapInline; // 提供默认值
        set { if (_wrapFormat != null) _wrapFormat.Type = value.EnumConvert(MsWord.WdWrapType.wdWrapInline); }
    }

    /// <inheritdoc/>
    public float DistanceTop
    {
        get => _wrapFormat?.DistanceTop ?? 0.0f;
        set { if (_wrapFormat != null) _wrapFormat.DistanceTop = value; }
    }

    /// <inheritdoc/>
    public float DistanceBottom
    {
        get => _wrapFormat?.DistanceBottom ?? 0.0f;
        set { if (_wrapFormat != null) _wrapFormat.DistanceBottom = value; }
    }

    /// <inheritdoc/>
    public float DistanceLeft
    {
        get => _wrapFormat?.DistanceLeft ?? 0.0f;
        set { if (_wrapFormat != null) _wrapFormat.DistanceLeft = value; }
    }

    /// <inheritdoc/>
    public float DistanceRight
    {
        get => _wrapFormat?.DistanceRight ?? 0.0f;
        set { if (_wrapFormat != null) _wrapFormat.DistanceRight = value; }
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 WrapFormat 封装类使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                if (_wrapFormat != null)
                {
                    Marshal.ReleaseComObject(_wrapFormat);
                    _wrapFormat = null;
                }
            }
            _disposedValue = true;
        }
    }

    /// <summary>
    /// 释放 WrapFormat 封装类使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}