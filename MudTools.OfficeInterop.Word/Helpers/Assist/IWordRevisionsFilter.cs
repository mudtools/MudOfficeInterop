//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word修订过滤器的接口，用于控制Word文档中修订的显示方式
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordRevisionsFilter : IOfficeObject<IWordRevisionsFilter>, IDisposable
{
    /// <summary>
    /// 获取或设置修订视图，用于控制文档中显示的修订类型
    /// </summary>
    WdRevisionsView View { get; set; }

    /// <summary>
    /// 获取或设置修订标记，用于控制修订在文档中的显示方式
    /// </summary>
    WdRevisionsMarkup Markup { get; set; }

    /// <summary>
    /// 获取与当前过滤器关联的审阅者集合
    /// </summary>
    IWordReviewers Reviewers { get; }

    /// <summary>
    /// 切换是否显示所有审阅者的修订
    /// </summary>
    void ToggleShowAllReviewers();
}