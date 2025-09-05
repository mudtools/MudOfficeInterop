//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.FileSearch 的二次封装实现类。
/// 提供安全访问文件搜索功能的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeFileSearch : IOfficeFileSearch
{
    private MsCore.FileSearch _fileSearch;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 FileSearch 对象。
    /// </summary>
    /// <param name="fileSearch">原始的 COM FileSearch 对象。</param>
    internal OfficeFileSearch(MsCore.FileSearch fileSearch)
    {
        _fileSearch = fileSearch ?? throw new ArgumentNullException(nameof(fileSearch));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IOfficeFoundFiles? FoundFiles
    {
        get
        {
            if (_fileSearch?.FoundFiles != null)
            {
                return new OfficeFoundFiles(_fileSearch.FoundFiles);
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeSearchScopes? SearchScopes
    {
        get
        {
            if (_fileSearch?.SearchScopes != null)
            {
                return new OfficeSearchScopes(_fileSearch.SearchScopes);
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficePropertyTests? PropertyTests
    {
        get
        {
            if (_fileSearch?.PropertyTests != null)
            {
                return new OfficePropertyTests(_fileSearch.PropertyTests);
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeSearchFolders? SearchFolders
    {
        get
        {
            if (_fileSearch?.SearchFolders != null)
            {
                return new OfficeSearchFolders(_fileSearch.SearchFolders);
            }
            return null;
        }
    }

    /// <inheritdoc/>
    public IOfficeFileTypes? FileTypes
    {
        get
        {
            if (_fileSearch?.FileTypes != null)
            {
                return new OfficeFileTypes(_fileSearch.FileTypes);
            }
            return null;
        }
    }


    /// <inheritdoc/>
    public MsoLastModified? LastModified
    {
        get => _fileSearch?.LastModified != null ? (MsoLastModified)(int)_fileSearch?.LastModified : MsoLastModified.msoLastModifiedToday;
        set
        {
            if (_fileSearch != null) _fileSearch.LastModified = (MsCore.MsoLastModified)(int)value;
        }
    }

    /// <inheritdoc/>
    public string FileName
    {
        get => _fileSearch?.FileName ?? string.Empty;
        set
        {
            if (_fileSearch != null)
                _fileSearch.FileName = value;
        }
    }

    /// <inheritdoc/>
    public bool MatchAllWordForms
    {
        get => _fileSearch?.MatchAllWordForms ?? false;
        set
        {
            if (_fileSearch != null)
                _fileSearch.MatchAllWordForms = value;
        }
    }

    /// <inheritdoc/>
    public bool MatchTextExactly
    {
        get => _fileSearch?.MatchTextExactly ?? false;
        set
        {
            if (_fileSearch != null)
                _fileSearch.MatchTextExactly = value;
        }
    }

    /// <inheritdoc/>
    public bool SearchSubFolders
    {
        get => _fileSearch?.SearchSubFolders ?? false;
        set
        {
            if (_fileSearch != null)
                _fileSearch.SearchSubFolders = value;
        }
    }

    public string LookIn
    {
        get => _fileSearch?.LookIn ?? string.Empty;
        set
        {
            if (_fileSearch != null)
                _fileSearch.LookIn = value;
        }
    }

    public string TextOrProperty
    {
        get => _fileSearch?.TextOrProperty ?? string.Empty;
        set
        {
            if (_fileSearch != null)
                _fileSearch.TextOrProperty = value;
        }
    }

    /// <inheritdoc/>
    public int FoundFilesCount => _fileSearch?.FoundFiles?.Count ?? 0;

    /// <inheritdoc/>
    public int PropertyTestsCount => _fileSearch?.PropertyTests?.Count ?? 0;


    #endregion

    #region 方法实现

    public void RefreshScopes()
    {
        _fileSearch?.RefreshScopes();
    }

    /// <inheritdoc/>
    public int Execute(MsoSortBy sortBy = MsoSortBy.msoSortByFileName,
                      MsoSortOrder sortOrder = MsoSortOrder.msoSortOrderAscending, bool alwaysAccurate = true)
    {
        if (_fileSearch == null)
            return 0;

        try
        {
            return _fileSearch.Execute((MsCore.MsoSortBy)(int)sortBy, (MsCore.MsoSortOrder)(int)sortOrder, alwaysAccurate);
        }
        catch
        {
            return 0;
        }
    }

    /// <inheritdoc/>
    public void NewSearch()
    {
        _fileSearch?.NewSearch();
    }
    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _fileSearch != null)
        {
            Marshal.ReleaseComObject(_fileSearch);
            _fileSearch = null;
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