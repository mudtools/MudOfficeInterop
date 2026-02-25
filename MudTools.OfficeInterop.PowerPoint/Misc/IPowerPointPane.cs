//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 文档窗口中的窗格。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPane : IOfficeObject<IPowerPointPane, MsPowerPoint.Pane>, IDisposable
{
    /// <summary>
    /// 获取此窗格的父对象。
    /// </summary>
    /// <value>表示此窗格父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 激活此窗格。
    /// </summary>
    void Activate();

    /// <summary>
    /// 获取一个值，指示此窗格是否为活动窗格。
    /// </summary>
    /// <value>指示是否为活动窗格的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Active { get; }

    /// <summary>
    /// 获取创建此窗格的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此窗格的视图类型。
    /// </summary>
    /// <value>表示视图类型的 <see cref="PpViewType"/> 枚举值。</value>
    PpViewType ViewType { get; }
}