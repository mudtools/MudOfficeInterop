//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.SearchScopes 的二次封装实现类。
/// 提供安全访问搜索范围集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeSearchScopes : IOfficeSearchScopes
{
    private MsCore.SearchScopes _searchScopes;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 SearchScopes 对象。
    /// </summary>
    /// <param name="searchScopes">原始的 COM SearchScopes 对象。</param>
    internal OfficeSearchScopes(MsCore.SearchScopes searchScopes)
    {
        _searchScopes = searchScopes ?? throw new ArgumentNullException(nameof(searchScopes));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _searchScopes?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeSearchScope this[int index]
    {
        get
        {
            if (_searchScopes == null || index < 1 || index > Count)
                return null;

            try
            {
                var scope = _searchScopes[index];
                return scope != null ? new OfficeSearchScope(scope) : null;
            }
            catch
            {
                return null;
            }
        }
    }
    #endregion

    #region IEnumerable<IOfficeSearchScope> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeSearchScope> GetEnumerator()
    {
        if (_searchScopes == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var scope = _searchScopes[i];
            if (scope != null)
                yield return new OfficeSearchScope(scope);
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

        if (disposing && _searchScopes != null)
        {
            try
            {
                Marshal.ReleaseComObject(_searchScopes);
            }
            catch
            {
                // 忽略释放异常
            }
            _searchScopes = null;
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