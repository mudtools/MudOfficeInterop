

namespace MudTools.OfficeInterop.Imps;


/// <summary>
/// IOfficeUserPermission 接口的内部实现类。
/// </summary>
internal class OfficeUserPermission : IOfficeUserPermission
{
    private MsCore.UserPermission? _userPermission;
    private bool _disposedValue;

    /// <summary>
    /// 初始化一个新的 OfficeUserPermission 实例。
    /// </summary>
    /// <param name="userPermission">原始的 Microsoft.Office.Core.UserPermission COM 对象。</param>
    /// <exception cref="ArgumentNullException">如果 <paramref name="userPermission"/> 为 null，则抛出此异常。</exception>
    internal OfficeUserPermission(MsCore.UserPermission userPermission)
    {
        _userPermission = userPermission ?? throw new ArgumentNullException(nameof(userPermission));
        _disposedValue = false;
    }

    /// <inheritdoc />
    public string UserId => _userPermission?.UserId ?? string.Empty;

    /// <inheritdoc />
    public DateTime ExpirationDate
    {
        get
        {
            if (_userPermission == null) return DateTime.MinValue;
            try
            {
                double oaDate = _userPermission.ExpirationDate.ConvertToDouble();
                if (oaDate == 0) // 0 通常表示没有到期日期
                    return DateTime.MinValue;
                return DateTime.FromOADate(oaDate);
            }
            catch (ArgumentException)
            {
                // 如果 OLE 日期无效，返回 MinValue
                return DateTime.MinValue;
            }
        }
        set
        {
            if (_userPermission != null)
            {
                double oaDate = value == DateTime.MinValue ? 0 : value.ToOADate();
                _userPermission.ExpirationDate = oaDate;
            }
        }
    }

    /// <inheritdoc />
    public int Permission => _userPermission?.Permission ?? 0;

    /// <inheritdoc />
    public void Remove()
    {
        try
        {
            _userPermission?.Remove();
            Marshal.ReleaseComObject(_userPermission);
            // 调用 Remove 后，COM 对象可能已失效，我们将其置空并标记为已释放
            _disposedValue = true;
            _userPermission = null;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法删除用户权限。该权限可能已被移除或文档状态已更改。", ex);
        }
    }

    /// <summary>
    /// 释放由 <see cref="OfficeUserPermission"/> 使用的资源。
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing && _userPermission != null)
            {
                Marshal.ReleaseComObject(_userPermission);
                _userPermission = null;
            }
            _disposedValue = true;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}