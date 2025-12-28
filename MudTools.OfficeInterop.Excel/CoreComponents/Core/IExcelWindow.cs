//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel 窗口接口，用于操作 Excel 窗口实例
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelWindow : IExcelCommonWindow, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 激活指定窗口，然后将其发送到窗口Z顺序的后面。
    /// </summary>
    /// <returns>操作结果。</returns>
    object? ActivateNext();

    /// <summary>
    /// 激活指定窗口，然后激活窗口Z顺序后面的窗口。
    /// </summary>
    /// <returns>操作结果。</returns>
    object? ActivatePrevious();

    /// <summary>
    /// 获取表示活动窗口（最顶层的窗口）或指定窗口中活动单元格的Range对象。如果窗口不显示工作表，此属性会失败。
    /// </summary>
    IExcelRange? ActiveCell { get; }

    /// <summary>
    /// 获取表示活动图表（嵌入图表或图表工作表）的Chart对象。当没有活动图表时，此属性返回null。
    /// </summary>
    IExcelChart? ActiveChart { get; }

    /// <summary>
    /// 获取表示窗口中活动窗格的Pane对象。
    /// </summary>
    IExcelPane? ActivePane { get; }

    /// <summary>
    /// 获取表示活动工作簿或指定窗口或工作簿中的活动工作表（位于顶部的）的对象。如果没有活动工作表，则返回null。
    /// </summary>
    object? ActiveSheet { get; }

    /// <summary>
    /// 关闭窗口。
    /// </summary>
    /// <param name="saveChanges">可选。如果没有对工作簿的更改，则忽略此参数。如果有更改且工作簿出现在其他打开的窗口中，也忽略此参数。如果有更改但工作簿未出现在任何其他打开的窗口中，此参数指定是否保存更改。</param>
    /// <param name="filename">可选。将更改保存到此文件名下。</param>
    /// <param name="routeWorkbook">可选。如果工作簿不需要路由到下一个收件人，则忽略此参数。否则，Microsoft Excel根据以下值路由工作簿：True表示发送给下一个收件人，False表示不发送，省略则显示对话框询问用户是否发送。</param>
    /// <returns>关闭操作的结果。</returns>
    bool? Close(bool? saveChanges = null, string? filename = null, bool? routeWorkbook = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示窗口是否显示公式；False表示窗口显示值。
    /// </summary>
    bool DisplayFormulas { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示网格线。
    /// </summary>
    bool DisplayGridlines { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示行号和列标；False表示不显示标题。
    /// </summary>
    bool DisplayHeadings { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示水平滚动条。
    /// </summary>
    bool DisplayHorizontalScrollBar { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示大纲符号。
    /// </summary>
    bool DisplayOutline { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示垂直滚动条。
    /// </summary>
    bool DisplayVerticalScrollBar { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示工作簿选项卡。
    /// </summary>
    bool DisplayWorkbookTabs { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示零值。
    /// </summary>
    bool DisplayZeros { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示窗口是否可以调整大小。
    /// </summary>
    bool EnableResize { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示拆分窗格是否被冻结。
    /// </summary>
    bool FreezePanes { get; set; }

    /// <summary>
    /// 获取或设置网格线颜色作为RGB值。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color GridlineColor { get; set; }

    /// <summary>
    /// 获取或设置网格线颜色作为当前调色板的索引或XlColorIndex常量。
    /// </summary>
    XlColorIndex GridlineColorIndex { get; set; }

    /// <summary>
    /// 获取对象在类似对象集合中的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 按页滚动窗口内容。
    /// </summary>
    /// <param name="down">可选。向下滚动内容的页数。</param>
    /// <param name="up">可选。向上滚动内容的页数。</param>
    /// <param name="toRight">可选。向右滚动内容的页数。</param>
    /// <param name="toLeft">可选。向左滚动内容的页数。</param>
    /// <returns>滚动操作的结果。</returns>
    object? LargeScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null);

    /// <summary>
    /// 为指定窗口创建新窗口或副本。
    /// </summary>
    /// <returns>新创建的Window对象。</returns>
    IExcelWindow? NewWindow();

    /// <summary>
    /// 获取或设置每当激活窗口时运行的过程的名称。
    /// </summary>
    string OnWindow { get; set; }

    /// <summary>
    /// 获取表示指定窗口中所有窗格的Panes集合。
    /// </summary>
    IExcelPanes? Panes { get; }

    /// <summary>
    /// 显示对象的打印预览。
    /// </summary>
    /// <param name="enableChanges">在预览期间启用更改。</param>
    /// <returns>打印预览操作的结果。</returns>
    object? PrintPreview(bool? enableChanges = null);

    /// <summary>
    /// 获取表示指定工作表中选定单元格的Range对象，即使工作表上有图形对象处于活动状态或已选中。
    /// </summary>
    IExcelRange? RangeSelection { get; }

    /// <summary>
    /// 获取或设置窗格或窗口中可见的最左侧列号。
    /// </summary>
    int ScrollColumn { get; set; }

    /// <summary>
    /// 获取或设置出现在窗格或窗口顶部的行号。
    /// </summary>
    int ScrollRow { get; set; }

    /// <summary>
    /// 在窗口底部滚动工作簿选项卡。
    /// </summary>
    /// <param name="sheets">可选。要滚动的表数。使用正数向前滚动，使用负数向后滚动，或使用0（零）表示不滚动。</param>
    /// <param name="position">可选。使用xlFirst滚动到第一个工作表，或使用xlLast滚动到最后一个工作表。</param>
    /// <returns>滚动操作的结果。</returns>
    object? ScrollWorkbookTabs(int? sheets = null, Constants? position = null);

    /// <summary>
    /// 获取表示指定窗口中所有选定工作表的Sheets集合。
    /// </summary>
    IExcelSheets? SelectedSheets { get; }

    /// <summary>
    /// 获取指定窗口中的选定对象。
    /// </summary>
    object Selection { get; }

    /// <summary>
    /// 按行或列滚动窗口内容。
    /// </summary>
    /// <param name="down">可选。向下滚动内容的行数。</param>
    /// <param name="up">可选。向上滚动内容的行数。</param>
    /// <param name="toRight">可选。向右滚动内容的列数。</param>
    /// <param name="toLeft">可选。向左滚动内容的列数。</param>
    /// <returns>滚动操作的结果。</returns>
    object? SmallScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null);

    /// <summary>
    /// 获取一个布尔值，表示窗口是否已拆分。
    /// </summary>
    bool Split { get; set; }

    /// <summary>
    /// 获取或设置窗口拆分为窗格的列号（拆分线左侧的列数）。
    /// </summary>
    int SplitColumn { get; set; }

    /// <summary>
    /// 获取或设置水平窗口拆分的位置（以磅为单位）。
    /// </summary>
    double SplitHorizontal { get; set; }

    /// <summary>
    /// 获取或设置窗口拆分为窗格的行号（拆分上方的行数）。
    /// </summary>
    int SplitRow { get; set; }

    /// <summary>
    /// 获取或设置垂直窗口拆分的位置（以磅为单位）。
    /// </summary>
    double SplitVertical { get; set; }

    /// <summary>
    /// 获取或设置工作簿选项卡区域的宽度与窗口水平滚动条宽度的比率（介于0（零）和1之间的数字；默认值为0.6）。
    /// </summary>
    double TabRatio { get; set; }


    /// <summary>
    /// 获取或设置窗口类型。
    /// </summary>
    XlWindowType Type { get; }

    /// <summary>
    /// 获取窗口在应用程序窗口区域中可以占据的最大高度（以磅为单位）。
    /// </summary>
    double UsableHeight { get; }

    /// <summary>
    /// 获取窗口在应用程序窗口区域中可以占据的最大宽度（以磅为单位）。
    /// </summary>
    double UsableWidth { get; }

    /// <summary>
    /// 获取表示窗口中或窗格中可见单元格范围的Range对象。如果列或行部分可见，则包含在范围内。
    /// </summary>
    IExcelRange? VisibleRange { get; }

    /// <summary>
    /// 获取窗口编号。例如，名为"Book1.xls:2"的窗口其窗口编号为2。大多数窗口的窗口编号为1。
    /// </summary>
    int WindowNumber { get; }

    /// <summary>
    /// 获取或设置窗口状态。
    /// </summary>
    XlWindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置窗口的显示大小作为百分比（100等于正常大小，200等于双倍大小，依此类推）。
    /// </summary>
    object Zoom { get; set; }

    /// <summary>
    /// 获取或设置窗口中显示的视图。
    /// </summary>
    XlWindowView View { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示指定窗口是否从右到左显示而不是从左到右显示。False表示对象从左到右显示。
    /// </summary>
    bool DisplayRightToLeft { get; set; }

    /// <summary>
    /// 将水平测量从点（文档坐标）转换为屏幕像素（屏幕坐标）。
    /// </summary>
    /// <param name="points">必需。从文档窗口顶部从左开始的水平点数。</param>
    /// <returns>像素坐标。</returns>
    int? PointsToScreenPixelsX(int points);

    /// <summary>
    /// 将垂直测量从点（文档坐标）转换为屏幕像素（屏幕坐标）。
    /// </summary>
    /// <param name="points">必需。从文档窗口左侧从顶部开始的垂直点数。</param>
    /// <returns>像素坐标。</returns>
    int? PointsToScreenPixelsY(int points);

    /// <summary>
    /// 获取位于指定屏幕坐标对上的Shape或Range对象。如果在指定坐标处没有形状，此方法返回null。
    /// </summary>
    /// <param name="x">必需。表示从屏幕左边缘水平距离的值（以像素为单位），从顶部开始。</param>
    /// <param name="y">必需。表示从屏幕顶部垂直距离的值（以像素为单位），从左侧开始。</param>
    /// <returns>位于指定坐标的对象。</returns>
    [MethodName("RangeFromPoint"), ValueConvert]
    IExcelRange? RangeFromPoint(int x, int y);

    /// <summary>
    /// 获取位于指定屏幕坐标对上的Shape或Range对象。如果在指定坐标处没有形状，此方法返回null。
    /// </summary>
    /// <param name="x">必需。表示从屏幕左边缘水平距离的值（以像素为单位），从顶部开始。</param>
    /// <param name="y">必需。表示从屏幕顶部垂直距离的值（以像素为单位），从左侧开始。</param>
    /// <returns>位于指定坐标的对象。</returns>
    [MethodName("RangeFromPoint"), ValueConvert]
    IExcelShape? ShapeFromPoint(int x, int y);

    /// <summary>
    /// 滚动文档窗口，使指定矩形区域的内容显示在文档窗口或窗格的左上角或右下角。
    /// </summary>
    /// <param name="left">必需。矩形水平位置（以磅为单位），从文档窗口或窗格的左边缘开始。</param>
    /// <param name="top">必需。矩形垂直位置（以磅为单位），从文档窗口或窗格的顶部开始。</param>
    /// <param name="width">必需。矩形的宽度（以磅为单位）。</param>
    /// <param name="height">必需。矩形的高度（以磅为单位）。</param>
    /// <param name="start">可选。True使矩形的左上角出现在文档窗口或窗格的左上角。False使矩形的右下角出现在文档窗口或窗格的右下角。默认值为True。</param>
    void ScrollIntoView(int left, int top, int width, int height, bool? start = null);

    /// <summary>
    /// 获取指定窗口的SheetViews对象。
    /// </summary>
    IExcelSheetViews? SheetViews { get; }

    /// <summary>
    /// 获取表示指定窗口中活动工作表视图的对象。
    /// </summary>
    object ActiveSheetView { get; }

    /// <summary>
    /// 打印对象。
    /// </summary>
    /// <param name="from">可选。开始打印的页码。如果省略此参数，则从开头开始打印。</param>
    /// <param name="to">可选。要打印的最后一页的页码。如果省略此参数，则打印到最后一页。</param>
    /// <param name="copies">可选。要打印的副本数。如果省略此参数，则打印一份。</param>
    /// <param name="preview">可选。True表示Excel在打印对象之前调用打印预览。False（或省略）表示立即打印对象。</param>
    /// <param name="activePrinter">可选。设置活动打印机的名称。</param>
    /// <param name="printToFile">可选。True表示打印到文件。如果未指定PrToFileName，则Excel会提示用户输入输出文件的名称。</param>
    /// <param name="collate">可选。True表示整理多份副本。</param>
    /// <param name="prToFileName">可选。如果PrintToFile设置为True，则此参数指定要打印到的文件的名称。</param>
    /// <returns>打印操作的结果。</returns>
    object? PrintOut(int? from = null, int? to = null, int? copies = null,
                    bool? preview = null, string? activePrinter = null,
                    bool? printToFile = null, bool? collate = null,
                    bool? prToFileName = null);

    /// <summary>
    /// 获取或设置一个布尔值，表示是否为指定窗口显示标尺。
    /// </summary>
    bool DisplayRuler { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示指定窗口中当前是否显示日期分组的自动筛选。
    /// </summary>
    bool AutoFilterDateGrouping { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，表示是否显示空白。
    /// </summary>
    bool DisplayWhitespace { get; set; }

    /// <summary>
    /// 获取窗口的句柄。
    /// </summary>
    int Hwnd { get; }
}