//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using MudTools.OfficeInterop.Word.Imps;

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 表示文件搜索操作的封装实现类。
/// </summary>
internal class WordFileSearch : IWordFileSearch
{
    private MsCore.FileSearch _fileSearch;
    private bool _disposedValue;

    /// <summary>
    /// 初始化 <see cref="WordFileSearch"/> 类的新实例。
    /// </summary>
    /// <param name="fileSearch">要封装的原始 COM FileSearch 对象。</param>
    internal WordFileSearch(MsCore.FileSearch fileSearch)
    {
        _fileSearch = fileSearch ?? throw new ArgumentNullException(nameof(fileSearch));
        _disposedValue = false;
    }

    #region 基本属性实现 (Basic Properties Implementation)
    /// <inheritdoc/>
    public int Creator => _fileSearch?.Creator ?? 0;

    #endregion

    #region 文件搜索属性实现 (File Search Properties Implementation)

    /// <inheritdoc/>
    public string TextOrProperty
    {
        get => _fileSearch?.TextOrProperty ?? string.Empty;
        set { if (_fileSearch != null) _fileSearch.TextOrProperty = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public string FileName
    {
        get => _fileSearch?.FileName ?? string.Empty;
        set { if (_fileSearch != null) _fileSearch.FileName = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public bool SearchSubFolders
    {
        get => _fileSearch?.SearchSubFolders ?? false;
        set { if (_fileSearch != null) _fileSearch.SearchSubFolders = value; }
    }

    /// <inheritdoc/>
    public bool MatchTextExactly
    {
        get => _fileSearch?.MatchTextExactly ?? true; // 默认通常包含文件
        set { if (_fileSearch != null) _fileSearch.MatchTextExactly = value; }
    }

    /// <inheritdoc/>
    public bool MatchAllWordForms
    {
        get => _fileSearch?.MatchAllWordForms ?? false; // 默认通常不包含文件夹
        set { if (_fileSearch != null) _fileSearch.MatchAllWordForms = value; }
    }

    /// <inheritdoc/>
    public MsoFileType FileType
    {
        get => _fileSearch?.FileType ?? MsoFileType.msoFileTypeAllFiles;
        set { if (_fileSearch != null) _fileSearch.FileType = value; }
    }

    /// <inheritdoc/>
    public IWordFoundFiles FoundFiles => _fileSearch?.FoundFiles != null ? new OfficeFoundFiles(_fileSearch.FoundFiles) : null;

    /// <inheritdoc/>
    public string LookIn
    {
        get => _fileSearch?.LookIn ?? string.Empty;
        set { if (_fileSearch != null) _fileSearch.LookIn = value ?? string.Empty; }
    }

    /// <inheritdoc/>
    public int LastModified => _fileSearch?.LastModified ?? 0;

    #endregion

    #region 文件搜索方法实现 (File Search Methods Implementation)

    /// <inheritdoc/>
    public int Execute()
    {
        if (_fileSearch == null) return 0;
        try
        {
            // Execute 方法返回找到的文件数
            return _fileSearch.Execute();
        }
        catch (COMException ex)
        {
            System.Diagnostics.Debug.WriteLine($"File search execution failed: {ex.Message}");
            return 0;
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        _fileSearch?.Reset();
    }

    /// <inheritdoc/>
    public MsWord.MsoFeatureInstall ShowSearchDialog()
    {
        if (_fileSearch == null) return MsWord.MsoFeatureInstall.msoFeatureInstallNone;
        try
        {
            // ShowSearchDialog 方法返回用户操作结果
            return _fileSearch.ShowSearchDialog();
        }
        catch (COMException ex)
        {
            System.Diagnostics.Debug.WriteLine($"Showing file search dialog failed: {ex.Message}");
            return MsWord.MsoFeatureInstall.msoFeatureInstallNone; // 返回默认值
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放由 <see cref="WordFileSearch"/> 使用的非托管资源，并选择性地释放托管资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
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

    /// <summary>
    /// 释放由 <see cref="WordFileSearch"/> 使用的所有资源。
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}