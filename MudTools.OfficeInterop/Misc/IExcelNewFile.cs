//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// NewFile 对象表示在多个 Microsoft Office 应用程序中可用的"新建项"任务窗格上列出的项目。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeNewFile : IOfficeObject<IOfficeNewFile, MsCore.NewFile>, IDisposable
{
    /// <summary>
    /// 获取一个表示对象的容器应用程序的 Application 对象。
    /// </summary>
    object? Application { get; }

    /// <summary>
    /// 获取一个指示创建指定对象的应用程序的 32 位整数。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 向"新建项"任务窗格添加新项目。
    /// </summary>
    /// <param name="fileName">要添加到任务窗格文件列表的文件名称。</param>
    /// <param name="section">要添加文件的区域。可以是任何 MsoFileNewSection 常量。</param>
    /// <param name="displayName">在任务窗格中显示的文本。</param>
    /// <param name="action">用户单击项目时采取的操作。可以是任何 MsoFileNewAction 常量。</param>
    /// <returns>如果添加成功，则为 true；否则为 false。</returns>
    bool? Add(string fileName, MsoFileNewSection? section = null, string? displayName = null, string? action = null);

    /// <summary>
    /// 从"新建项"任务窗格中删除项目。
    /// </summary>
    /// <param name="fileName">文件引用的名称。</param>
    /// <param name="section">任务窗格中存在文件引用的区域。可以是任何 MsoFileNewSection 常量。</param>
    /// <param name="displayName">文件引用的显示文本。</param>
    /// <param name="action">用户单击项目时采取的操作。可以是任何 MsoFileNewAction 常量。</param>
    /// <returns>如果删除成功，则为 true；否则为 false。</returns>
    bool? Remove(string fileName, MsoFileNewSection? section = null, string? displayName = null, string? action = null);
}