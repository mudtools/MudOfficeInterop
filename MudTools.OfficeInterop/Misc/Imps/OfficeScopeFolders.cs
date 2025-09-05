//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;
/// <summary>
/// 对 Microsoft.Office.Core.ScopeFolders 的二次封装实现类。
/// 提供安全访问搜索范围文件夹集合的方式，并管理 COM 对象生命周期。
/// </summary>
internal class OfficeScopeFolders : IOfficeScopeFolders
{
    private MsCore.ScopeFolders _scopeFolders;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，初始化封装的 ScopeFolders 对象。
    /// </summary>
    /// <param name="scopeFolders">原始的 COM ScopeFolders 对象。</param>
    internal OfficeScopeFolders(MsCore.ScopeFolders scopeFolders)
    {
        _scopeFolders = scopeFolders ?? throw new ArgumentNullException(nameof(scopeFolders));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public int Count => _scopeFolders?.Count ?? 0;

    /// <inheritdoc/>
    public IOfficeScopeFolder this[int index]
    {
        get
        {
            if (_scopeFolders == null || index < 1 || index > Count)
                return null;

            try
            {
                var scopeFolder = _scopeFolders[index];
                return scopeFolder != null ? new OfficeScopeFolder(scopeFolder) : null;
            }
            catch
            {
                return null;
            }
        }
    }
    #endregion

    #region IEnumerable<IOfficeScopeFolder> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeScopeFolder> GetEnumerator()
    {
        if (_scopeFolders == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var scopeFolder = _scopeFolders[i];
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

        if (disposing && _scopeFolders != null)
        {
            try
            {
                Marshal.ReleaseComObject(_scopeFolders);
            }
            catch
            {
                // 忽略释放异常
            }
            _scopeFolders = null;
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