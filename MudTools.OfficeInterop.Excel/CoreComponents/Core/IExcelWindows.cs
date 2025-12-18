//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel 窗口集合接口，用于操作 Excel 窗口集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelWindows : IDisposable, IEnumerable<IExcelWindow>
{
    /// <summary>
    /// 获取窗口数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取父对象（通常是 Application）
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取所在的Application对象
    /// 对应 Windows.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 根据索引获取窗口（从1开始）
    /// </summary>
    /// <param name="index">窗口索引</param>
    /// <returns>窗口对象</returns>
    IExcelWindow? this[int index] { get; }

    /// <summary>
    /// 根据窗口标题获取窗口
    /// </summary>
    /// <param name="caption">窗口标题</param>
    /// <returns>窗口对象</returns>
    IExcelWindow? this[string caption] { get; }

    /// <summary>
    /// 对屏幕上的窗口进行排列。
    /// </summary>
    /// <param name="arrangeStyle">指定窗口在屏幕上的排列方式。</param>
    /// <param name="activeWorkbook">如果为 True，则只排列活动工作簿的可见窗口。 如果为 False，则排列所有窗口。 默认值为 False。</param>
    /// <param name="syncHorizontal">如果为 True，则在水平滚动时同步活动工作簿的窗口。 False 表示不同步窗口。 默认值为 False。</param>
    /// <param name="syncVertical">如果为 True，则在垂直滚动时同步活动工作簿的窗口。 False 表示不同步窗口。 默认值为 False。</param>
    /// <returns></returns>
    object? Arrange(
        XlArrangeStyle arrangeStyle = XlArrangeStyle.xlArrangeStyleTiled,
        bool? activeWorkbook = false,
        bool? syncHorizontal = false,
        bool? syncVertical = false);

    /// <summary>
    /// 可结束两个窗口的并排模式。
    /// </summary>
    /// <returns></returns>
    bool? BreakSideBySide();

    /// <summary>
    /// 以并排模式打开两个窗口。
    /// </summary>
    /// <param name="WindowName">要打开的窗口的名称。</param>
    /// <returns></returns>
    bool? CompareSideBySideWith(string? WindowName);

    /// <summary>
    /// 重置正在进行并排比较的两个工作表窗口的位置。
    /// </summary>
    void ResetPositionsSideBySide();
}