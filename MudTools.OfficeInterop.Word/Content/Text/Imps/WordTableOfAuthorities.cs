
using log4net;

namespace MudTools.OfficeInterop.Word.Imps;

internal class WordTableOfAuthorities : IWordTableOfAuthorities
{
    private static readonly ILog log = LogManager.GetLogger(typeof(WordTableOfAuthorities));
    private MsWord.TableOfAuthorities? _tableOfAuthorities;
    private bool _disposedValue;

    internal WordTableOfAuthorities(MsWord.TableOfAuthorities tableOfAuthorities)
    {
        _tableOfAuthorities = tableOfAuthorities ?? throw new ArgumentNullException(nameof(tableOfAuthorities));
        _disposedValue = false;
    }

    #region 属性实现

    public IWordApplication? Application => _tableOfAuthorities != null ? new WordApplication(_tableOfAuthorities.Application) : null;

    public object? Parent => _tableOfAuthorities?.Parent;

    public IWordRange? Range => _tableOfAuthorities?.Range != null ? new WordRange(_tableOfAuthorities.Range) : null;

    public bool Passim
    {
        get => _tableOfAuthorities?.Passim ?? false;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.Passim = value;
        }
    }

    public string? EntrySeparator
    {
        get => _tableOfAuthorities?.EntrySeparator;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.EntrySeparator = value;
        }
    }

    public string? PageRangeSeparator
    {
        get => _tableOfAuthorities?.PageRangeSeparator;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.PageRangeSeparator = value;
        }
    }

    public string? Bookmark
    {
        get => _tableOfAuthorities?.Bookmark;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.Bookmark = value;
        }
    }

    public int Category
    {
        get => _tableOfAuthorities?.Category ?? 0;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.Category = value;
        }
    }

    public bool KeepEntryFormatting
    {
        get => _tableOfAuthorities?.KeepEntryFormatting ?? false;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.KeepEntryFormatting = value;
        }
    }

    public string Separator
    {
        get => _tableOfAuthorities?.Separator ?? string.Empty;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.Separator = value;
        }
    }

    public string IncludeSequenceName
    {
        get => _tableOfAuthorities?.IncludeSequenceName ?? string.Empty;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.IncludeSequenceName = value;
        }
    }

    public bool IncludeCategoryHeader
    {
        get => _tableOfAuthorities?.IncludeCategoryHeader ?? false;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.IncludeCategoryHeader = value;
        }
    }

    public string PageNumberSeparator
    {
        get => _tableOfAuthorities?.PageNumberSeparator ?? string.Empty;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.PageNumberSeparator = value;
        }
    }

    public WdTabLeader TabLeader
    {
        get => _tableOfAuthorities != null ? _tableOfAuthorities.TabLeader.EnumConvert(WdTabLeader.wdTabLeaderLines) : WdTabLeader.wdTabLeaderLines;
        set
        {
            if (_tableOfAuthorities != null)
                _tableOfAuthorities.TabLeader = value.EnumConvert(MsWord.WdTabLeader.wdTabLeaderLines);
        }
    }


    #endregion

    #region 方法实现

    public void Update()
    {
        try
        {
            _tableOfAuthorities?.Update();
        }
        catch (Exception ex)
        {
            log.Error("更新引文目录失败。", ex);
            throw new InvalidOperationException("更新引文目录失败。", ex);
        }
    }

    public void Delete()
    {
        try
        {
            _tableOfAuthorities?.Delete();
        }
        catch (Exception ex)
        {
            log.Error("删除引文目录失败。", ex);
            throw new InvalidOperationException("删除引文目录失败。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _tableOfAuthorities != null)
        {
            Marshal.ReleaseComObject(_tableOfAuthorities);
            _tableOfAuthorities = null;
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}