//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 应用程序中所有任务窗格的对象集合。
/// <para>注：使用 Application.TaskPanes 属性可返回 TaskPanes 集合。</para>
/// <para>注：使用 TaskPanes(index)（其中 index 是任务窗格类型或索引号）可返回单个 TaskPane 对象。</para>
/// <para>注：此接口基于对 Word 对象模型和 Office 应用程序中 TaskPanes 的普遍理解实现，因为官方 SDK 文档信息有限。</para>
/// </summary>
public interface IWordTaskPanes : IEnumerable<IWordTaskPane>, IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中的任务窗格数量。
    /// </summary>
    int Count { get; }

    #endregion

    #region 集合索引器 (Collection Indexer)

    /// <summary>
    /// 通过索引（任务窗格类型或索引号）获取单个任务窗格。
    /// </summary>
    /// <param name="index">任务窗格类型 (<see cref="WdTaskPanes"/>) 或索引号（整数）。</param>
    /// <returns>指定的任务窗格对象。</returns>
    IWordTaskPane this[WdTaskPanes index] { get; }

    #endregion
}