//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示活动文档的权限设置。此接口是对 Microsoft.Office.Core.Permission COM 对象的封装。
/// 使用此对象可以限制对活动文档的访问权限，并返回或设置特定的权限设置 [[1]], [[14]].
/// </summary>
public interface IOfficePermission : IEnumerable<IOfficeUserPermission>, IDisposable
{
    /// <summary>
    /// 获取一个值，该值指示是否已为活动文档启用了权限限制。
    /// </summary>
    /// <value>
    /// 如果为活动文档启用了权限限制，则为 <see langword="true"/>；否则为 <see langword="false"/>。
    /// </value>
    bool Enabled { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示当用户尝试在受支持的应用程序（如 Word）中打开具有受限权限的文档但没有客户端权限管理支持时，
    /// 是否允许其在受信任的浏览器中查看该文档 [[31]], [[35]].
    /// </summary>
    /// <value>
    /// 如果启用受信任的浏览器支持，则为 <see langword="true"/>；否则为 <see langword="false"/>。默认值为 <see langword="false"/>。
    /// </value>
    bool EnableTrustedBrowser { get; set; }

    /// <summary>
    /// 获取权限集合中的项数（即具有权限的用户数）。
    /// 如果未在活动文档上启用权限，则返回 0。启用权限后，此值至少为 1（包含文档作者）[[22]], [[24]].
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取一个值，该值指示上次打开活动文档时是否应用了权限策略 [[13]].
    /// </summary>
    /// <value>
    /// 如果应用了权限策略，则为 <see langword="true"/>；否则为 <see langword="false"/>。
    /// </value>
    bool PermissionFromPolicy { get; }

    /// <summary>
    /// 获取应用于活动文档的权限策略的名称 [[26]].
    /// </summary>
    string PolicyName { get; }

    /// <summary>
    /// 获取应用于活动文档的权限策略的描述 [[25]].
    /// </summary>
    string PolicyDescription { get; }

    /// <summary>
    /// 获取或设置活动文档的作者。作者始终对文档拥有非过期的所有者权限 [[76]].
    /// </summary>
    string DocumentAuthor { get; set; }

    /// <summary>
    /// 获取或设置需要额外权限的用户的联系信息（文件或网站 URL，或电子邮件地址）[[8]].
    /// </summary>
    string RequestPermissionURL { get; set; }

    /// <summary>
    /// 获取指定用户的 <see cref="IOfficeUserPermission"/> 对象。
    /// 索引可以是用户的电子邮件地址（字符串）或集合中的位置（从 1 开始的整数）[[72]], [[77]].
    /// </summary>
    /// <param name="index">用户的电子邮件地址或从 1 开始的索引位置。</param>
    /// <returns>表示指定用户的权限的 <see cref="IOfficeUserPermission"/> 对象；如果指定的索引不存在，则返回 <see langword="null"/>。</returns>
    IOfficeUserPermission? this[int index] { get; }

    /// <summary>
    /// 为指定用户创建一组针对活动文档的新权限 [[42]], [[48]].
    /// </summary>
    /// <param name="userId">要授予权限的用户的电子邮件地址，格式为 user@domain.com。</param>
    /// <param name="permission">要授予指定用户的权限，可以是 <see cref="MsCore.MsoPermission"/> 值的一个或多个组合。</param>
    /// <param name="expirationDate">权限的可选到期日期。如果未指定，则权限永不过期。</param>
    /// <returns>表示新创建的用户权限的 <see cref="IOfficeUserPermission"/> 对象。</returns>
    /// <exception cref="ArgumentException">当 <paramref name="userId"/> 为空或格式无效时抛出。</exception>
    /// <exception cref="Exception">当 COM 调用失败时抛出。</exception>
    IOfficeUserPermission Add(string userId, object? permission = null, DateTime? expirationDate = null);

    /// <summary>
    /// 将指定的权限策略应用于活动文档。
    /// </summary>
    /// <param name="policyFileName">要应用的权限策略文件的名称。</param>
    void ApplyPolicy(string policyFileName);

    /// <summary>
    /// 从活动文档的权限集合中移除所有 <see cref="IOfficeUserPermission"/> 对象，并禁用对活动文档的限制 [[60]], [[64]].
    /// </summary>
    /// <exception cref="Exception">当 COM 调用失败时抛出。</exception>
    void RemoveAll();
}