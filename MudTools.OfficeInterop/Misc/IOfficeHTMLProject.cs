//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office应用程序中的HTML项目，用于处理HTML文档和脚本编辑
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeHTMLProject : IOfficeObject<IOfficeHTMLProject, MsCore.HTMLProject>, IDisposable
{
    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 返回 Microsoft.Office.Core.HTMLProject 对象的当前状态。
    /// </summary>
    MsoHTMLProjectState State { get; }

    /// <summary>
    /// 返回包含在指定 HTML 项目中的 Microsoft.Office.Core.HTMLProjectItems 集合。
    /// </summary>
    IOfficeHTMLProjectItems? HTMLProjectItems { get; }

    /// <summary>
    /// 在 Microsoft 脚本编辑器中刷新指定的 HTML 项目。
    /// </summary>
    /// <param name="refresh">可选 Boolean。如果要刷新文档，则为 True；如果不刷新文档，则为 False。</param>
    void RefreshProject(bool refresh = true);

    /// <summary>
    /// 在 Microsoft Office 主机应用程序中刷新指定的 HTML 项目。
    /// </summary>
    /// <param name="refresh">可选 Boolean。如果要保存所有更改，则为 True；如果要忽略所有更改，则为 False。</param>
    void RefreshDocument(bool refresh = true);

    /// <summary>
    /// 在 Microsoft 脚本编辑器中打开指定的 HTML 项目或 HTML 项目项，使用可选的 Microsoft.Office.Core.MsoHTMLProjectOpen 常量指定的视图之一。
    /// </summary>
    /// <param name="openKind">可选 Microsoft.Office.Core.MsoHTMLProjectOpen。打开指定项目或项目项的视图。</param>
    void Open(MsoHTMLProjectOpen openKind = MsoHTMLProjectOpen.msoHTMLProjectOpenSourceView);
}