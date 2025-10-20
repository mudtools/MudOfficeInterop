
using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordTablesOfContents : IWordTablesOfContents
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordTablesOfContents));
    private readonly DisposableList _disposableList = new();
    private MsWord.TablesOfContents? _tablesOfContents;
    private bool _disposedValue;

    internal WordTablesOfContents(MsWord.TablesOfContents tablesOfContents)
    {
        _tablesOfContents = tablesOfContents ?? throw new ArgumentNullException(nameof(tablesOfContents));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _tablesOfContents != null ? new WordApplication(_tablesOfContents.Application) : null;

    public object? Parent => _tablesOfContents?.Parent;

    public int Count => _tablesOfContents?.Count ?? 0;

    public IWordTableOfContents? this[int index]
    {
        get
        {
            if (_tablesOfContents == null || index < 1 || index > Count) return null;
            try
            {
                var comToc = _tablesOfContents[index];
                var result = comToc != null ? new WordTableOfContents(comToc) : null;
                if (result != null)
                    _disposableList.Add(result);
                return result;
            }
            catch (COMException ce)
            {
                log.Error($"根据索引 {index} 检索目录对象失败: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    public IWordTableOfContents Add(
        IWordRange range,
        bool useHeadingStyles = true,
        int upperHeadingLevel = 1,
        int lowerHeadingLevel = 3,
        bool useFields = false,
        string? tableId = null)
    {
        if (range == null)
            throw new ArgumentNullException(nameof(range));

        if (_tablesOfContents == null)
            throw new InvalidOperationException("目录集合不可用。");

        try
        {
            // 将二次封装的 Range 转换回原始 COM 对象
            var comRange = (range as WordRange)?._range;
            if (comRange == null)
                throw new ArgumentException("提供的范围对象无效。", nameof(range));

            var comToc = _tablesOfContents.Add(
                Range: comRange,
                UseHeadingStyles: useHeadingStyles,
                UpperHeadingLevel: upperHeadingLevel,
                LowerHeadingLevel: lowerHeadingLevel,
                UseFields: useFields,
                TableID: tableId ?? string.Empty
            );

            var tocWrapper = new WordTableOfContents(comToc);
            _disposableList.Add(tocWrapper);
            return tocWrapper;
        }
        catch (Exception ex)
        {
            log.Error("向文档添加新目录失败。", ex);
            throw new InvalidOperationException("向文档添加新目录失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tablesOfContents != null)
        {
            Marshal.ReleaseComObject(_tablesOfContents);
            _disposableList.Dispose();
            _tablesOfContents = null;
        }
        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion

    #region IEnumerable<IWordTableOfContents> 实现

    public IEnumerator<IWordTableOfContents> GetEnumerator()
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