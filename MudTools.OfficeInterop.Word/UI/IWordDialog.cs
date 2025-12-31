//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示一个内置对话框。
/// <para>注：Dialog 对象是 Dialogs 集合的成员。Dialogs 集合包含 Microsoft Word 中的所有内置对话框。</para>
/// <para>注：无法创建新的内置对话框，或添加到 Dialogs 集合。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordDialog : IOfficeObject<IWordDialog>, IDisposable
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

    #region 对话框属性 (Dialog Properties)

    /// <summary>
    /// 获取内置 Microsoft Word 对话框的工具栏控件 ID。此为只读属性。
    /// </summary>
    int CommandBarId { get; }

    /// <summary>
    /// 获取在指定的内置对话框中显示的过程的名称。
    /// </summary>
    string CommandName { get; }

    /// <summary>
    /// 获取或设置显示指定对话框时的活动选项卡。
    /// </summary>
    WdWordDialogTab DefaultTab { get; set; }

    /// <summary>
    /// 获取内置 Microsoft Word 对话框的类型。
    /// </summary>
    WdWordDialog Type { get; }

    #endregion

    #region 对话框方法 (Dialog Methods)

    /// <summary>
    /// 显示指定的内置 Microsoft Word 对话框。
    /// </summary>
    /// <param name="timeout">指定对话框显示的秒数。达到超时值后，对话框自动关闭。</param>
    /// <returns>如果用户单击“确定”则返回 true，如果用户单击“取消”则返回 false。</returns>
    [ValueConvert]
    bool? Display(float? timeout = null);

    /// <summary>
    /// 应用 Microsoft Word 对话框的当前设置。
    /// </summary>
    void Execute();

    /// <summary>
    /// 显示并执行在指定的内置 Microsoft Word 对话框中启动的操作。
    /// </summary>
    /// <param name="timeout">指定对话框显示的秒数。达到超时值后，对话框自动关闭。</param>
    /// <returns>如果用户单击“确定”则返回 true，如果用户单击“取消”则返回 false。</returns>
    [ValueConvert]
    bool? Show(float? timeout = null);

    /// <summary>
    /// 更新内置 Microsoft Word 对话框中显示的值。
    /// </summary>
    void Update();

    #endregion
}
