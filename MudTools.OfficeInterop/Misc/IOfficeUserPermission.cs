

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示与活动文档关联的单个用户的权限集。
/// 此对象将一组权限、一个用户（由其电子邮件地址标识）和一个可选的到期日期绑定在一起。
/// 此接口是对 Microsoft.Office.Core.UserPermission COM 对象的封装。
/// </summary>
public interface IOfficeUserPermission : IDisposable
{
    /// <summary>
    /// 获取与该权限集关联的用户的电子邮件地址。
    /// </summary>
    /// <value>用户的电子邮件地址字符串。</value>
    string UserId { get; }

    /// <summary>
    /// 获取或设置分配给此用户的权限的可选到期日期。
    /// 如果权限没有到期日期，则返回 <see cref="DateTime.MinValue"/>。
    /// 设置为 <see cref="DateTime.MinValue"/> 或早于当前日期的日期可以移除到期日期或使权限立即失效。
    /// </summary>
    /// <value>表示权限到期日期的 <see cref="DateTime"/> 对象。</value>
    DateTime ExpirationDate { get; set; }

    /// <summary>
    /// 获取一个值，该值表示分配给此用户的权限类型。
    /// 该值是 <see cref="MsCore.MsoPermission"/> 枚举的按位组合。
    /// </summary>
    /// <value>一个整数，代表权限的按位组合。</value>
    int Permission { get; }

    /// <summary>
    /// 从活动文档的权限集合中删除此用户及其关联的权限集。
    /// 调用此方法后，该接口实例不应再被使用。
    /// </summary>
    /// <exception cref="InvalidOperationException">如果底层 COM 对象已被释放或删除，则抛出此异常。</exception>
    void Remove();
}