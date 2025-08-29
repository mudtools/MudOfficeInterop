namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TabStops 的实现类。
/// </summary>
internal class WordTabStops : IWordTabStops
{
    private MsWord.TabStops _tabStops;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="tabStops">原始 COM TabStops 对象。</param>
    internal WordTabStops(MsWord.TabStops tabStops)
    {
        _tabStops = tabStops ?? throw new ArgumentNullException(nameof(tabStops));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _tabStops?.Count ?? 0;

    /// <inheritdoc/>
    public IWordTabStop this[int index]
    {
        get
        {
            if (_tabStops == null || index < 1 || index > Count)
                return null;

            var tabStop = _tabStops[index];
            return new WordTabStop(tabStop);
        }
    }

    /// <inheritdoc/>
    public IWordTabStop this[float position]
    {
        get
        {
            if (_tabStops == null)
                return null;

            var tabStop = _tabStops[position];
            return tabStop != null ? new WordTabStop(tabStop) : null;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public IWordTabStop Add(float position,
        WdTabAlignment alignment = WdTabAlignment.wdAlignTabLeft,
        WdTabLeader leader = WdTabLeader.wdTabLeaderSpaces)
    {
        if (_tabStops == null)
            throw new ObjectDisposedException(nameof(WordTabStops));

        try
        {
            var newTabStop = _tabStops.Add(position, (MsWord.WdTabAlignment)(int)alignment, (MsWord.WdTabLeader)(int)leader);
            return new WordTabStop(newTabStop);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加制表符到位置 {position}。", ex);
        }
    }

    /// <inheritdoc/>
    public IWordTabStop Find(float position)
    {
        if (_tabStops == null)
            return null;

        try
        {
            var tabStop = _tabStops[position];
            return tabStop != null ? new WordTabStop(tabStop) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <inheritdoc/>
    public void Clear(float position)
    {
        if (_tabStops == null)
            return;

        _tabStops[position]?.Clear();
    }

    /// <inheritdoc/>
    public void ClearAll()
    {
        _tabStops?.ClearAll();
    }

    /// <inheritdoc/>
    public List<float> GetPositions()
    {
        var positions = new List<float>();

        if (_tabStops == null)
            return positions;

        for (int i = 1; i <= Count; i++)
        {
            var tabStop = _tabStops[i];
            if (tabStop != null)
                positions.Add(tabStop.Position);
        }

        return positions;
    }

    #endregion

    #region IEnumerable<IWordTabStop> 实现

    /// <inheritdoc/>
    public IEnumerator<IWordTabStop> GetEnumerator()
    {
        if (_tabStops == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var tabStop = _tabStops[i];
            if (tabStop != null)
                yield return new WordTabStop(tabStop);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
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

        if (disposing && _tabStops != null)
        {
            Marshal.ReleaseComObject(_tabStops);
            _tabStops = null;
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