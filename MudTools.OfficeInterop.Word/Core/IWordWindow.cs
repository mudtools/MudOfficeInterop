//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示一个 Word 文档窗口。
/// 封装了 Microsoft.Office.Interop.Word.Window 对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordWindow : IOfficeObject<IWordWindow, MsWord.Window>, IDisposable
{
    #region 属性

    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置是否显示垂直滚动条。
    /// </summary>
    bool? DisplayVerticalScrollBar { get; set; }


    /// <summary>
    /// 获取或设置是否显示水平滚动条。
    /// </summary>
    bool? DisplayHorizontalScrollBar { get; set; }

    /// <summary>
    /// 获取窗口的文档窗格集合。
    /// </summary>
    IWordPanes? Panes { get; }

    /// <summary>
    /// 获取或设置窗口是否处于激活状态。
    /// </summary>
    bool? Active { get; }

    /// <summary>
    /// 获取Excel应用程序窗口的句柄
    /// </summary>
    int? Hwnd { get; }

    /// <summary>
    /// 获取窗口的文档。
    /// </summary>
    IWordDocument? Document { get; }

    /// <summary>
    /// 获取窗口的文档视图设置。
    /// </summary>
    IWordView? View { get; }

    /// <summary>
    /// 获取下一个文档窗口。
    /// </summary>
    IWordWindow? Next { get; }

    /// <summary>
    /// 获取上一个文档窗口。
    /// </summary>
    IWordWindow? Previous { get; }

    /// <summary>
    /// 获取窗口的垂直位置（以磅为单位）。
    /// </summary>
    int? VerticalPercentScrolled { get; set; }

    /// <summary>
    /// 获取窗口的水平位置（以磅为单位）。
    /// </summary>
    int? HorizontalPercentScrolled { get; set; }

    /// <summary>
    /// 获取窗口的高度（以磅为单位）。
    /// </summary>
    int? Height { get; set; }

    /// <summary>
    /// 获取窗口的宽度（以磅为单位）。
    /// </summary>
    int? Width { get; set; }

    /// <summary>
    /// 获取窗口的标题栏文本。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取窗口的索引号。
    /// </summary>
    int? Index { get; }

    /// <summary>
    /// 获取或设置窗口的左坐标。
    /// </summary>
    int? Left { get; set; }

    /// <summary>
    /// 获取或设置窗口的顶坐标。
    /// </summary>
    int? Top { get; set; }

    /// <summary>
    /// 获取或设置窗口是否可见。
    /// </summary>
    bool? Visible { get; set; }

    /// <summary>
    /// 获取窗口的类型。
    /// </summary>
    WdWindowType Type { get; }

    /// <summary>
    /// 获取或设置窗口是否最大化。
    /// </summary>
    WdWindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置窗口的分隔位置。
    /// </summary>
    int? SplitVertical { get; set; }

    /// <summary>
    /// 获取窗口的可用宽度（以磅为单位）。
    /// </summary>
    int UsableWidth { get; }

    /// <summary>
    /// 获取窗口的可用高度（以磅为单位）。
    /// </summary>
    int UsableHeight { get; }

    /// <summary>
    /// 获取或设置信封是否可见。
    /// </summary>
    bool EnvelopeVisible { get; set; }

    /// <summary>
    /// 获取或设置是否显示右侧标尺。
    /// </summary>
    bool DisplayRightRuler { get; set; }

    /// <summary>
    /// 获取或设置是否显示左侧滚动条。
    /// </summary>
    bool DisplayLeftScrollBar { get; set; }

    /// <summary>
    /// 获取或设置缩略图是否可见。
    /// </summary>
    bool Thumbnails { get; set; }

    /// <summary>
    /// 获取或设置源文档的显示方式。
    /// </summary>
    WdShowSourceDocuments ShowSourceDocuments { get; set; }

    /// <summary>
    /// 获取当前活动窗格。
    /// </summary>
    IWordPane? ActivePane { get; }

    /// <summary>
    /// 获取当前选择内容。
    /// </summary>
    IWordSelection? Selection { get; }

    /// <summary>
    /// 获取或设置输入法编辑器 (IME) 模式。
    /// </summary>
    WdIMEMode IMEMode { get; set; }

    /// <summary>
    /// 获取窗口编号。
    /// </summary>
    int WindowNumber { get; }

    /// <summary>
    /// 获取或设置文档地图的百分比宽度。
    /// </summary>
    int DocumentMapPercentWidth { get; set; }

    /// <summary>
    /// 获取或设置窗口是否已拆分。
    /// </summary>
    bool Split { get; set; }

    /// <summary>
    /// 获取或设置是否显示标尺。
    /// </summary>
    bool DisplayRulers { get; set; }

    /// <summary>
    /// 获取或设置是否显示垂直标尺。
    /// </summary>
    bool DisplayVerticalRuler { get; set; }

    /// <summary>
    /// 获取或设置样式区域的宽度。
    /// </summary>
    float StyleAreaWidth { get; set; }

    /// <summary>
    /// 获取或设置是否显示屏幕提示。
    /// </summary>
    bool DisplayScreenTips { get; set; }

    /// <summary>
    /// 获取或设置文档地图是否可见。
    /// </summary>
    bool DocumentMap { get; set; }

    #endregion // 属性

    #region 方法

    /// <summary>
    /// 激活指定的窗口。
    /// </summary>
    void Activate();

    /// <summary>
    /// 关闭指定的窗口。
    /// </summary>
    /// <param name="saveChanges">指定保存更改的方式。</param>
    /// <param name="routeDocument">如果为 true，则将文档路由给下一个收件人。</param>
    void Close(WdSaveOptions saveChanges = WdSaveOptions.wdPromptToSaveChanges, bool routeDocument = false);


    /// <summary>
    /// 创建一个与指定窗口具有相同文档的新窗口。
    /// </summary>
    IWordWindow? NewWindow();


    /// <summary>
    /// 将焦点设置到指定的窗口。
    /// </summary>
    void SetFocus();

    /// <summary>
    /// 返回文档中指定屏幕坐标处的区域<see cref="IWordRange"/>对象。
    /// </summary>
    /// <param name="x">屏幕坐标中的X轴位置（以像素为单位）。</param>
    /// <param name="y">屏幕坐标中的Y轴位置（以像素为单位）。</param>
    /// <returns>返回指定点处的范围对象，如果点不在文档窗口中则返回null。</returns>
    [ValueConvert, MethodName("RangeFromPoint")]
    IWordRange? RangeFromPoint(int x, int y);

    /// <summary>
    /// 返回文档中指定屏幕坐标处的区域<see cref="IWordShape"/>对象。
    /// </summary>
    /// <param name="x">屏幕坐标中的X轴位置（以像素为单位）。</param>
    /// <param name="y">屏幕坐标中的Y轴位置（以像素为单位）。</param>
    /// <returns>返回指定点处的范围对象，如果点不在文档窗口中则返回null。</returns>
    [ValueConvert, MethodName("RangeFromPoint")]
    IWordShape? ShapeFromPoint(int x, int y);

    /// <summary>
    /// 返回指定对象在屏幕上的位置和大小信息（以像素为单位）。
    /// </summary>
    /// <param name="ScreenPixelsLeft">返回对象在屏幕上的水平位置。</param>
    /// <param name="ScreenPixelsTop">返回对象在屏幕上的垂直位置。</param>
    /// <param name="ScreenPixelsWidth">返回对象的宽度。</param>
    /// <param name="ScreenPixelsHeight">返回对象的高度。</param>
    /// <param name="obj">要获取位置信息的对象。</param>
    void GetPoint(out int ScreenPixelsLeft, out int ScreenPixelsTop, out int ScreenPixelsWidth, out int ScreenPixelsHeight, object obj);

    /// <summary>
    /// 将窗口滚动到下一页。
    /// </summary>
    void LargeScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null);

    /// <summary>
    /// 将窗口滚动到指定位置。
    /// </summary>
    void ScrollIntoView(IWordRange range, bool? scrollToTopOfRange = null);

    /// <summary>
    /// 将窗口滚动到指定位置。
    /// </summary>
    void ScrollIntoView(IWordShape range, bool? scrollToTopOfRange = null);

    /// <summary>
    /// 将窗口滚动到下一页。
    /// </summary>
    void PageScroll(int? pages = null, int? lines = null);

    /// <summary>
    /// 将窗口滚动到下一页。
    /// </summary>
    void SmallScroll(int? down = null, int? up = null, int? toRight = null, int? toLeft = null);

    /// <summary>
    /// 切换是否显示所有审阅者批注的可见性
    /// </summary>
    void ToggleShowAllReviewers();

    /// <summary>
    /// 切换功能区的可见性
    /// </summary>
    void ToggleRibbon();

    /// <summary>
    /// 打印文档
    /// </summary>
    /// <param name="background">如果为 true，则在后台打印文档</param>
    /// <param name="append">如果为 true，则将文档追加到输出文件末尾</param>
    /// <param name="range">指定打印范围</param>
    /// <param name="outputFileName">输出文件的名称</param>
    /// <param name="from">起始页码</param>
    /// <param name="to">结束页码</param>
    /// <param name="item">指定要打印的项目类型</param>
    /// <param name="copies">打印份数</param>
    /// <param name="pages">指定要打印的页码范围，例如"1-5, 7, 9-12"</param>
    /// <param name="pageType">指定打印的页面类型</param>
    /// <param name="printToFile">如果为 true，则将输出发送到文件</param>
    /// <param name="collate">如果为 true，则对副本进行排序</param>
    /// <param name="activePrinterMacGX">在 Macintosh 上使用，如果为 true，则使用活动打印机设置</param>
    /// <param name="manualDuplexPrint">如果为 true，则手动双面打印</param>
    /// <param name="printZoomColumn">指定打印时每页的列数</param>
    /// <param name="printZoomRow">指定打印时每页的行数</param>
    /// <param name="printZoomPaperWidth">指定打印时的纸张宽度（磅）</param>
    /// <param name="printZoomPaperHeight">指定打印时的纸张高度（磅）</param>
    void PrintOut(
        bool? background = null,
        bool? append = null,
        WdPrintOutRange? range = null,
        string? outputFileName = null,
        int? from = null,
        int? to = null,
        WdPrintOutItem? item = null,
        int? copies = null,
        string? pages = null,
        WdPrintOutPages? pageType = null,
        bool? printToFile = null,
        bool? collate = null,
        bool? activePrinterMacGX = null,
        bool? manualDuplexPrint = null,
        string? printZoomColumn = null,
        string? printZoomRow = null,
        double? printZoomPaperWidth = null,
        double? printZoomPaperHeight = null);

    #endregion // 方法
}