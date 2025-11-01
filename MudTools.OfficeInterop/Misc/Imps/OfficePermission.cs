//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Imps;

/// <summary>
/// IOfficePermission 接口的内部实现类。
/// </summary>
internal class OfficePermission : IOfficePermission
{
    private MsCore.Permission? _permission;
    private bool _disposedValue;

    internal OfficePermission(MsCore.Permission permission)
    {
        _permission = permission ?? throw new ArgumentNullException(nameof(permission));
        _disposedValue = false;
    }

    public bool Enabled
    {
        get => _permission?.Enabled ?? false;
        set
        {
            if (_permission != null)
                _permission.Enabled = value;
        }

    }

    public int Count => _permission?.Count ?? 0;

    public string DocumentAuthor
    {
        get => _permission?.DocumentAuthor ?? string.Empty;
        set
        {
            if (_permission != null)
                _permission.DocumentAuthor = value;
        }
    }

    public string RequestPermissionURL
    {
        get => _permission?.RequestPermissionURL ?? string.Empty;
        set
        {
            if (_permission != null)
                _permission.RequestPermissionURL = value;
        }
    }

    public bool PermissionFromPolicy => _permission?.PermissionFromPolicy ?? false;

    public string PolicyName => _permission?.PolicyName ?? string.Empty;

    public string PolicyDescription => _permission?.PolicyDescription ?? string.Empty;

    public bool EnableTrustedBrowser
    {
        get => _permission?.EnableTrustedBrowser ?? false;
        set
        {
            if (_permission != null)
                _permission.EnableTrustedBrowser = value;
        }
    }

    /// <inheritdoc/>
    public IOfficeUserPermission? this[int index]
    {
        get
        {
            if (_permission == null || index < 1 || index > Count)
                return null;

            try
            {
                var field = _permission[index];
                return field != null ? new OfficeUserPermission(field) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    public void ApplyPolicy(string policyFileName)
    {
        if (string.IsNullOrEmpty(policyFileName))
            throw new ArgumentException("策略文件名不能为空。", nameof(policyFileName));

        try
        {
            _permission?.ApplyPolicy(policyFileName);
        }
        catch (Exception ex)
        {
            throw new Exception("ApplyPolicy 方法调用失败！", ex);
        }
    }

    public IOfficeUserPermission Add(string userId, object? permission = null, DateTime? expirationDate = null)
    {
        if (string.IsNullOrEmpty(userId))
            throw new ArgumentException("用户ID不能为空。", nameof(userId));

        try
        {
            if (_permission != null)
            {
                object comExpirationDate = expirationDate?.ToOADate() ?? Type.Missing;
                var userPermisson = _permission.Add(userId, permission ?? Type.Missing, comExpirationDate);
                return new OfficeUserPermission(userPermisson);
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new Exception("Add 方法调用失败！", ex);
        }
    }

    public void RemoveAll()
    {
        try
        {
            _permission?.RemoveAll();
        }
        catch (Exception ex)
        {
            throw new Exception("RemoveAll 方法调用失败！", ex);
        }
    }

    #region IEnumerable<IOfficePickerField> 实现

    /// <inheritdoc/>
    public IEnumerator<IOfficeUserPermission> GetEnumerator()
    {
        if (_permission == null)
            yield break;

        for (int i = 1; i <= Count; i++)
        {
            var field = _permission[i];
            if (field != null)
                yield return new OfficeUserPermission(field);
        }
    }

    /// <inheritdoc/>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing && _permission != null)
            {
                Marshal.ReleaseComObject(_permission);
                _permission = null;
            }
            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}