//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.SearchFolders 的二次封装实现类。
/// 提供安全访问搜索文件夹集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeSearchFolders : IOfficeSearchFolders
{
    private static readonly ILog log = LogManager.GetLogger(typeof(OfficePropertyTests));
    private MsCore.SearchFolders _searchFolders;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 SearchFolders 对象。
    /// </summary>
    /// <param name="searchFolders">原始的 COM SearchFolders 对象。</param>
    internal OfficeSearchFolders(MsCore.SearchFolders searchFolders)
    {
        _searchFolders = searchFolders ?? throw new ArgumentNullException(nameof(searchFolders));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _searchFolders?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeScopeFolder this[int index]
    {
        get
        {
            if (_searchFolders == null || index < 1 || index > Count)
                return null;

            try
            {
                var scopeFolder = _searchFolders[index];
                return scopeFolder != null ? new OfficeScopeFolder(scopeFolder) : null;
            }
            catch (COMException ce)
            {
                log.Error($"Failed to retrieve object based on index: {ce.Message}", ce);
                return null;
            }
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Add(IOfficeScopeFolder scopeFolder)
    {
        if (_searchFolders == null || scopeFolder == null)
            return;

        _searchFolders.Add(((OfficeScopeFolder)scopeFolder)._scopeFolder);
    }

    /// <inheritdoc/>
    public void Remove(int index)
    {
        if (_searchFolders == null || index < 1 || index > Count)
            return;

        _searchFolders.Remove(index);
    }

    /// <inheritdoc/>
    public void Clear()
    {
        while (Count > 0)
        {
            _searchFolders.Remove(1);
        }
    }

    #endregion

    #region IEnumerable<IOfficeScopeFolder> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeScopeFolder> GetEnumerator()
    {
        if (_searchFolders == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var scopeFolder = _searchFolders[i];
            if (scopeFolder != null)
                yield return new OfficeScopeFolder(scopeFolder);
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
    /// 释放资源的核心方法。
    /// </summary>
    /// <param name="disposing">是否由 Dispose() 调用。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _searchFolders != null)
        {
            Marshal.ReleaseComObject(_searchFolders);
            _searchFolders = null;
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