namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TabStop 的实现类。
/// </summary>
internal class WordTabStop : IWordTabStop
{
    private MsWord.TabStop _tabStop;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="tabStop">原始 COM TabStop 对象。</param>
    internal WordTabStop(MsWord.TabStop tabStop)
    {
        _tabStop = tabStop ?? throw new ArgumentNullException(nameof(tabStop));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public float Position
    {
        get => _tabStop?.Position ?? 0f;
        set
        {
            if (_tabStop != null)
                _tabStop.Position = value;
        }
    }

    /// <inheritdoc/>
    public WdTabAlignment Alignment
    {
        get => (WdTabAlignment)(int)_tabStop?.Alignment;
        set
        {
            if (_tabStop != null)
                _tabStop.Alignment = (MsWord.WdTabAlignment)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdTabLeader Leader
    {
        get => (WdTabLeader)(int)_tabStop?.Leader;
        set
        {
            if (_tabStop != null)
                _tabStop.Leader = (MsWord.WdTabLeader)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool CustomTab => _tabStop.CustomTab;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Clear()
    {
        _tabStop?.Clear();
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tabStop != null)
        {
            Marshal.ReleaseComObject(_tabStop);
            _tabStop = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}