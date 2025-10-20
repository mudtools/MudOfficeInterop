
using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordTablesOfAuthorities : IWordTablesOfAuthorities
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordTablesOfAuthorities));
    private readonly DisposableList _disposableList = new();
    private MsWord.TablesOfAuthorities? _tablesOfAuthorities;
    private bool _disposedValue;

    internal WordTablesOfAuthorities(MsWord.TablesOfAuthorities tablesOfAuthorities)
    {
        _tablesOfAuthorities = tablesOfAuthorities ?? throw new ArgumentNullException(nameof(tablesOfAuthorities));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _tablesOfAuthorities != null ? new WordApplication(_tablesOfAuthorities.Application) : null;

    public object? Parent => _tablesOfAuthorities?.Parent;

    public int Count => _tablesOfAuthorities?.Count ?? 0;

    public IWordTableOfAuthorities? this[int index]
    {
        get
        {
            if (_tablesOfAuthorities == null || index < 1 || index > Count) return null;
            try
            {
                var comToa = _tablesOfAuthorities[index];
                var result = comToa != null ? new WordTableOfAuthorities(comToa) : null;
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                log.Error($"根据索引 {index} 检索引文目录对象失败: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    public IWordTableOfAuthorities Add(
        IWordRange range,
        int category = 0,
        string? bookmark = null,
        string? entrySeparator = null,
        string? pageRangeSeparator = null,
        bool passim = false)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        if (_tablesOfAuthorities == null)
            throw new InvalidOperationException("引文目录集合不可用。");

        try
        {
            var comRange = (range as WordRange)?._range;
            if (comRange == null)
                throw new ArgumentException("提供的范围对象无效。", nameof(range));

            var comToa = _tablesOfAuthorities.Add(
                Range: comRange,
                Category: category,
                Bookmark: bookmark ?? string.Empty,
                EntrySeparator: entrySeparator ?? string.Empty,
                PageRangeSeparator: pageRangeSeparator ?? string.Empty,
                Passim: passim
            );

            var toaWrapper = new WordTableOfAuthorities(comToa);
            _disposableList.Add(toaWrapper);
            return toaWrapper;
        }
        catch (Exception ex)
        {
            log.Error("向文档添加新引文目录失败。", ex);
            throw new InvalidOperationException("向文档添加新引文目录失败。", ex);
        }
    }

    public IWordField MarkCitation(
        IWordRange range,
        string entry,
        string? shortCitation = null,
        int category = 0)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));
        if (string.IsNullOrEmpty(entry))
            throw new ArgumentException("引文条目不能为空。", nameof(entry));

        if (_tablesOfAuthorities == null)
            throw new InvalidOperationException("无法标记引文，因为引文目录集合不可用。");

        try
        {
            var comRange = (range as WordRange)?._range;
            if (comRange == null)
                throw new ArgumentException("提供的范围对象无效。", nameof(range));

            var comField = _tablesOfAuthorities.MarkCitation(
                Range: comRange,
                ShortCitation: shortCitation ?? string.Empty,
                LongCitation: entry,
                Category: category
            );

            var fieldWrapper = new WordField(comField);
            // Note: Fields are often managed by the document, so adding to _disposableList is optional.
            return fieldWrapper;
        }
        catch (Exception ex)
        {
            log.Error("标记引文条目失败。", ex);
            throw new InvalidOperationException("标记引文条目失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tablesOfAuthorities != null)
        {
            Marshal.ReleaseComObject(_tablesOfAuthorities);
            _disposableList.Dispose();
            _tablesOfAuthorities = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordTableOfAuthorities> 实现

    public IEnumerator<IWordTableOfAuthorities> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
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