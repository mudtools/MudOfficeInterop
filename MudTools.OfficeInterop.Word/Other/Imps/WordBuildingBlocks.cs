
namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 对 <see cref="MsWord.BuildingBlocks"/> COM 集合的封装实现类。
/// 表示某一类型（如页眉）和类别（如“常规”）下的所有构建基块条目。
/// 提供安全的访问、添加、查询和资源管理机制。
/// </summary>
internal class WordBuildingBlocks : IWordBuildingBlocks
{
    private MsWord.BuildingBlocks _blocks;
    private bool _disposedValue;

    /// <summary>
    /// 使用指定的 COM BuildingBlocks 集合初始化封装实例。
    /// </summary>
    /// <param name="blocks">原始的 BuildingBlocks COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 blocks 为 null 时抛出。</exception>
    internal WordBuildingBlocks(MsWord.BuildingBlocks blocks)
    {
        _blocks = blocks ?? throw new ArgumentNullException(nameof(blocks));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc />
    public int Count
    {
        get
        {
            if (_disposedValue) throw new ObjectDisposedException(nameof(WordBuildingBlocks));

            try
            {
                return _blocks?.Count ?? 0;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法获取构建基块数量。", ex);
            }
        }
    }

    #endregion

    #region 索引器实现

    /// <inheritdoc />
    public IWordBuildingBlock? this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            var comBlock = _blocks.Item(index);
            if (comBlock == null) return null;

            return new WordBuildingBlock(comBlock);
        }
    }

    /// <inheritdoc />
    public IWordBuildingBlock? this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            var comBlock = _blocks.Item(name);
            if (comBlock == null) return null;

            return new WordBuildingBlock(comBlock);
        }
    }

    #endregion

    #region 方法实现

    public IWordBuildingBlock Add(string name, string value, WdDocPartInsertOptions insertOptions)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("名称不能为空。", nameof(name));
        value ??= string.Empty;

        try
        {
            // 获取当前活动文档的内容范围，用于设置临时内容
            var app = _blocks.Application;
            var doc = app.ActiveDocument;
            var range = doc.Content;
            range.Text = value;

            var newBlock = _blocks.Add(
                Name: name,
                Range: range,
                InsertOptions: (MsWord.WdDocPartInsertOptions)(int)insertOptions
            );

            // 清空临时内容
            range.Text = string.Empty;

            return new WordBuildingBlock(newBlock);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException($"无法添加构建基块 '{name}'。", ex);
        }
    }

    /// <inheritdoc />
    public bool Contains(string name)
    {
        if (_disposedValue || string.IsNullOrWhiteSpace(name)) return false;

        return _blocks.Item(name) != null;
    }

    /// <inheritdoc />
    public List<string> GetNames()
    {
        var names = new List<string>();

        if (_disposedValue || _blocks == null) return names;

        for (int i = 1; i <= Count; i++)
        {
            var block = _blocks.Item(i);
            if (block?.Name != null)
            {
                names.Add(block.Name.ToString());
            }
        }

        return names;
    }

    /// <inheritdoc />
    public void Clear()
    {
        if (_disposedValue || _blocks == null) return;

        try
        {
            // 倒序删除，防止索引错乱
            for (int i = Count; i >= 1; i--)
            {
                var block = _blocks.Item(i);
                block?.Delete();
            }
        }
        catch (COMException)
        {
            // 忽略整体异常
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

        if (disposing && _blocks != null)
        {
            Marshal.ReleaseComObject(_blocks);
            _blocks = null;
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
    public IEnumerator<IWordBuildingBlock> GetEnumerator()
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