//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 的文档窗口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointDocumentWindow : IOfficeObject<IPowerPointDocumentWindow, MsPowerPoint.DocumentWindow>, IDisposable
{
    /// <summary>
    /// 获取创建此文档窗口的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此文档窗口的父对象。
    /// </summary>
    /// <value>表示此文档窗口父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取文档窗口中当前选中的对象。
    /// </summary>
    /// <value>表示当前选中的 <see cref="IPowerPointSelection"/> 对象。</value>
    IPowerPointSelection? Selection { get; }

    /// <summary>
    /// 获取文档窗口的视图对象。
    /// </summary>
    /// <value>表示视图的 <see cref="IPowerPointView"/> 对象。</value>
    IPowerPointView? View { get; }

    /// <summary>
    /// 获取文档窗口中打开的演示文稿。
    /// </summary>
    /// <value>表示演示文稿的 <see cref="IPowerPointPresentation"/> 对象。</value>
    IPowerPointPresentation? Presentation { get; }

    /// <summary>
    /// 获取或设置文档窗口的视图类型。
    /// </summary>
    /// <value>表示视图类型的 <see cref="PpViewType"/> 枚举值。</value>
    PpViewType ViewType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文档窗口是否以黑白模式显示。
    /// </summary>
    /// <value>指示是否以黑白模式显示的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool BlackAndWhite { get; set; }

    /// <summary>
    /// 获取一个值，指示文档窗口是否为活动窗口。
    /// </summary>
    /// <value>指示是否为活动窗口的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Active { get; }

    /// <summary>
    /// 获取或设置文档窗口的状态。
    /// </summary>
    /// <value>表示窗口状态的 <see cref="PpWindowState"/> 枚举值。</value>
    PpWindowState WindowState { get; set; }

    /// <summary>
    /// 获取文档窗口的标题。
    /// </summary>
    /// <value>表示窗口标题的字符串。</value>
    string? Caption { get; }

    /// <summary>
    /// 获取或设置文档窗口的左边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示左边缘位置的浮点数。</value>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置文档窗口的上边缘位置（以磅为单位）。
    /// </summary>
    /// <value>表示上边缘位置的浮点数。</value>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置文档窗口的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示宽度的浮点数。</value>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置文档窗口的高度（以磅为单位）。
    /// </summary>
    /// <value>表示高度的浮点数。</value>
    float Height { get; set; }

    /// <summary>
    /// 调整文档窗口以适应页面。
    /// </summary>
    void FitToPage();

    /// <summary>
    /// 激活文档窗口。
    /// </summary>
    void Activate();

    /// <summary>
    /// 在文档窗口中大范围滚动。
    /// </summary>
    /// <param name="down">向下滚动的次数。</param>
    /// <param name="up">向上滚动的次数。</param>
    /// <param name="toRight">向右滚动的次数。</param>
    /// <param name="toLeft">向左滚动的次数。</param>
    void LargeScroll(int down = 1, int up = 0, int toRight = 0, int toLeft = 0);

    /// <summary>
    /// 在文档窗口中小范围滚动。
    /// </summary>
    /// <param name="down">向下滚动的次数。</param>
    /// <param name="up">向上滚动的次数。</param>
    /// <param name="toRight">向右滚动的次数。</param>
    /// <param name="toLeft">向左滚动的次数。</param>
    void SmallScroll(int down = 1, int up = 0, int toRight = 0, int toLeft = 0);

    /// <summary>
    /// 创建包含相同文档的新窗口。
    /// </summary>
    /// <returns>新创建的 <see cref="IPowerPointDocumentWindow"/> 对象。</returns>
    IPowerPointDocumentWindow? NewWindow();

    /// <summary>
    /// 关闭文档窗口。
    /// </summary>
    void Close();

    /// <summary>
    /// 获取文档窗口的窗口句柄。
    /// </summary>
    /// <value>表示窗口句柄的整数值。</value>
    int HWND { get; }

    /// <summary>
    /// 获取当前活动的窗格。
    /// </summary>
    /// <value>表示活动窗格的 <see cref="IPowerPointPane"/> 对象。</value>
    IPowerPointPane? ActivePane { get; }

    /// <summary>
    /// 获取文档窗口的窗格集合。
    /// </summary>
    /// <value>表示窗格集合的 <see cref="IPowerPointPanes"/> 对象。</value>
    IPowerPointPanes? Panes { get; }

    /// <summary>
    /// 获取或设置垂直拆分位置（以磅为单位）。
    /// </summary>
    /// <value>表示垂直拆分位置的整数值。</value>
    int SplitVertical { get; set; }

    /// <summary>
    /// 获取或设置水平拆分位置（以磅为单位）。
    /// </summary>
    /// <value>表示水平拆分位置的整数值。</value>
    int SplitHorizontal { get; set; }

    /// <summary>
    /// 返回位于指定屏幕坐标处的对象。
    /// </summary>
    /// <param name="x">屏幕 X 坐标。</param>
    /// <param name="y">屏幕 Y 坐标。</param>
    /// <returns>位于指定坐标处的对象。</returns>
    object? RangeFromPoint(int x, int y);

    /// <summary>
    /// 将水平距离从磅转换为屏幕像素。
    /// </summary>
    /// <param name="points">以磅为单位的水平距离。</param>
    /// <returns>转换后的屏幕像素值。</returns>
    int? PointsToScreenPixelsX(float points);

    /// <summary>
    /// 将垂直距离从磅转换为屏幕像素。
    /// </summary>
    /// <param name="points">以磅为单位的垂直距离。</param>
    /// <returns>转换后的屏幕像素值。</returns>
    int? PointsToScreenPixelsY(float points);

    /// <summary>
    /// 将指定区域滚动到视图中。
    /// </summary>
    /// <param name="left">区域的左边缘位置（以磅为单位）。</param>
    /// <param name="top">区域的上边缘位置（以磅为单位）。</param>
    /// <param name="width">区域的宽度（以磅为单位）。</param>
    /// <param name="height">区域的高度（以磅为单位）。</param>
    /// <param name="start">指示滚动开始位置的布尔值。</param>
    void ScrollIntoView(float left, float top, float width, float height, [ConvertTriState] bool start = true);

    /// <summary>
    /// 检查指定节是否展开。
    /// </summary>
    /// <param name="sectionIndex">节的索引。</param>
    /// <returns>如果节已展开，则返回 true；否则返回 false。</returns>
    bool? IsSectionExpanded(int sectionIndex);

    /// <summary>
    /// 展开或折叠指定节。
    /// </summary>
    /// <param name="sectionIndex">节的索引。</param>
    /// <param name="expand">如果为 true，则展开节；如果为 false，则折叠节。</param>
    void ExpandSection(int sectionIndex, bool expand);
}
