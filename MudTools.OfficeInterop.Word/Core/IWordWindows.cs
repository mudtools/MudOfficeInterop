//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 窗口集合接口
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordWindows : IDisposable, IEnumerable<IWordWindow?>
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取窗口数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象（通常是 Application）
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置是否启用并排窗口的同步滚动功能
    /// </summary>
    /// <value>
    /// <c>true</c> 启用并排窗口的同步滚动；<c>false</c> 禁用同步滚动
    /// </value>
    bool SyncScrollingSideBySide { get; set; }

    /// <summary>
    /// 根据索引获取窗口（从1开始）
    /// </summary>
    IWordWindow? this[int index] { get; }

    /// <summary>
    /// 根据窗口标题获取窗口
    /// </summary>
    IWordWindow? this[string caption] { get; }

    /// <summary>
    /// 添加一个新窗口，用于显示当前活动文档
    /// </summary>
    /// <param name="window">要添加的窗口对象，如果为null则创建新窗口</param>
    /// <returns>新添加的IWordWindow对象，如果失败则返回null</returns>
    IWordWindow? Add(IWordWindow? window = null);

    /// <summary>
    /// 排列所有打开的文档窗口
    /// </summary>
    /// <param name="arrangeStyle">窗口排列样式，如果为null则使用默认样式</param>
    void Arrange(WdArrangeStyle? arrangeStyle = null);

    /// <summary>
    /// 将指定文档与当前活动文档进行并排比较
    /// </summary>
    /// <param name="document">要与当前文档进行比较的Word文档，如果为null则使用当前活动文档</param>
    /// <returns>如果成功启动并排比较则返回true，否则返回false或null</returns>
    bool? CompareSideBySideWith(IWordDocument? document = null);

    /// <summary>
    /// 断开并排比较模式
    /// </summary>
    /// <returns>如果成功断开并排比较则返回true，否则返回false或null</returns>
    bool? BreakSideBySide();

    /// <summary>
    /// 重置并排比较窗口的位置
    /// </summary>
    void ResetPositionsSideBySide();
}