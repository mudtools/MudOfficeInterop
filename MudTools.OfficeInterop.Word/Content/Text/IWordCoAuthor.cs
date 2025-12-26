//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中的一个协作者，提供对协作者信息的访问
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordCoAuthor : IOfficeObject<IWordCoAuthor>, IDisposable
{
    /// <summary>
    /// 获取与此协作者关联的Word应用程序实例
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此协作者的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取协作者的显示名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取一个值，该值指示当前协作者是否是当前用户
    /// </summary>
    bool IsMe { get; }

    /// <summary>
    /// 获取与此协作者关联的锁定集合
    /// </summary>
    IWordCoAuthLocks Locks { get; }

    /// <summary>
    /// 获取协作者的电子邮件地址
    /// </summary>
    string EmailAddress { get; }
}