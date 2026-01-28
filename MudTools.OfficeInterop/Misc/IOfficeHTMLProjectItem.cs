//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Microsoft 脚本编辑器中项目资源管理器中的项目项分支的单个项目项。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeHTMLProjectItem : IOfficeObject<IOfficeHTMLProjectItem, MsCore.HTMLProjectItem>, IDisposable
{
    /// <summary>
    /// 获取指定对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 返回指定对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 确定指定的 HTML 项目项是否在 Microsoft 脚本编辑器中打开。
    /// </summary>
    bool IsOpen { get; }

    /// <summary>
    /// 获取或设置 HTML 编辑器中的 HTML 文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 从指定文件（磁盘上）的文本更新 Microsoft 脚本编辑器中的文本。
    /// </summary>
    /// <param name="fileName">必需 String。包含要加载文本的文本文件的完全限定路径。</param>
    void LoadFromFile(string fileName);

    /// <summary>
    /// 在 Microsoft 脚本编辑器中打开指定的 HTML 项目或 HTML 项目项，使用可选的 Microsoft.Office.Core.MsoHTMLProjectOpen 常量指定的视图之一。
    /// </summary>
    /// <param name="openKind">可选 Microsoft.Office.Core.MsoHTMLProjectOpen。打开指定项目或项目项的视图。</param>
    void Open(MsoHTMLProjectOpen openKind = MsoHTMLProjectOpen.msoHTMLProjectOpenSourceView);

    /// <summary>
    /// 使用新文件名保存指定的 HTML 项目项。
    /// </summary>
    /// <param name="fileName">必需 String。要保存 HTML 项目项的文件的完全限定路径。</param>
    void SaveCopyAs(string fileName);
}
