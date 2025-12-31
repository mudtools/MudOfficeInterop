//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 应用程序中的一个任务窗格。
/// <para>注：TaskPane 对象是 TaskPanes 集合的成员。</para>
/// <para>注：使用 TaskPanes(index) 可返回单个 TaskPane 对象，其中 index 是索引号。</para>
/// <para>注：此接口基于对 Word 对象模型和 Office 应用程序中 TaskPane 的普遍理解实现，因为官方 SDK 文档信息有限。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTaskPane : IOfficeObject<IWordTaskPane>, IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #endregion

    #region 任务窗格属性 (Task Pane Properties)
    /// <summary>
    /// 获取或设置一个值，该值指示任务窗格是否可见。
    /// </summary>
    bool Visible { get; set; }
    #endregion
}