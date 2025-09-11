//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel 窗口接口，用于操作 Excel 窗口实例
/// </summary>
public interface IExcelWindow : IExcelCommonWindow, IDisposable
{
    /// <summary>
    /// 获取所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取窗口状态（最大化、最小化、正常）
    /// </summary>
    XlWindowState WindowState { get; set; }

    /// <summary>
    /// 获取或设置窗口视图类型
    /// </summary>
    XlWindowView View { get; set; }

    /// <summary>
    /// 获取或设置显示比例（75 表示 75%、100 等于正常大小、200 等于两倍大小等）
    /// </summary>
    double Zoom { get; set; }

    /// <summary>
    /// 获取或设置是否冻结窗格
    /// </summary>
    bool FreezePanes { get; set; }

    /// <summary>
    /// 获取或设置拆分窗口的行位置
    /// </summary>
    int SplitRow { get; set; }

    /// <summary>
    /// 获取或设置拆分窗口的列位置
    /// </summary>
    int SplitColumn { get; set; }

    /// <summary>
    /// 获取是否已拆分窗口
    /// </summary>
    bool Split { get; set; }

    /// <summary>
    /// 获取或设置是否显示网格线
    /// </summary>
    bool DisplayGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否显示行列标题
    /// </summary>
    bool DisplayHeadings { get; set; }

    /// <summary>
    /// 获取或设置是否显示零值
    /// </summary>
    bool DisplayZeros { get; set; }

    /// <summary>
    /// 获取或设置是否从右到左显示
    /// </summary>
    bool DisplayRightToLeft { get; set; }
    /// <summary>
    /// 获取或设置是否显示公式
    /// </summary>
    bool DisplayFormulas { get; set; }

    /// <summary>
    /// 获取或设置是否显示水平滚动条
    /// </summary>
    bool DisplayHorizontalScrollBar { get; set; }

    /// <summary>
    /// 获取或设置是否显示垂直滚动条
    /// </summary>
    bool DisplayVerticalScrollBar { get; set; }

    /// <summary>
    /// 获取或设置是否显示工作表标签
    /// </summary>
    bool DisplayWorkbookTabs { get; set; }

    /// <summary>
    /// 获取或设置是否显示标尺
    /// </summary>
    bool DisplayRuler { get; set; }

    /// <summary>
    /// 获取或设置是否启用自动筛选器日期分组功能
    /// </summary>
    bool AutoFilterDateGrouping { get; set; }

    /// <summary>
    /// 获取或设置是否显示空白字符
    /// </summary>
    bool DisplayWhitespace { get; set; }

    /// <summary>
    /// 获取或设置当前垂直滚动位置（行号）
    /// </summary>
    int ScrollRow { get; set; }

    /// <summary>
    /// 获取或设置当前水平滚动位置（列号）
    /// </summary>
    int ScrollColumn { get; set; }


    /// <summary>
    /// 获取或设置是否显示分级显示符号
    /// </summary>
    bool DisplayOutline { get; set; }

    /// <summary>
    /// 获取或设置网格线颜色（RGB值）
    /// </summary>
    int GridlineColor { get; set; }

    /// <summary>
    /// 获取或设置网格线颜色索引
    /// </summary>
    XlColorIndex GridlineColorIndex { get; set; }

    /// <summary>
    /// 获取窗口句柄（HWND）
    /// </summary>
    int Hwnd { get; }

    /// <summary>
    /// 获取窗口中的窗格集合
    /// </summary>
    IExcelPanes Panes { get; }

    /// <summary>
    /// 获取当前选中的单元格区域
    /// </summary>
    IExcelRange? RangeSelection { get; }

    /// <summary>
    /// 获取当前选中的单元格区域
    /// </summary>
    IExcelRange? Selection { get; }

    /// <summary>
    /// 获取工作表视图集合
    /// </summary>
    IExcelSheetViews SheetViews { get; }

    /// <summary>
    /// 获取或设置水平拆分位置（像素）
    /// </summary>
    double SplitHorizontal { get; set; }

    /// <summary>
    /// 获取或设置垂直拆分位置（像素）
    /// </summary>
    double SplitVertical { get; set; }

    /// <summary>
    /// 获取或设置工作表标签区域占比
    /// </summary>
    double TabRatio { get; set; }

    /// <summary>
    /// 获取窗口类型（工作表/图表）
    /// </summary>
    XlWindowType Type { get; }

    /// <summary>
    /// 获取窗口可用高度（排除工具栏等）
    /// </summary>
    double UsableHeight { get; }

    /// <summary>
    /// 获取窗口可用宽度（排除工具栏等）
    /// </summary>
    double UsableWidth { get; }

    /// <summary>
    /// 获取当前窗口中可见的单元格区域
    /// </summary>
    IExcelRange? VisibleRange { get; }

    /// <summary>
    /// 根据指定的坐标点获取对应的单元格区域
    /// </summary>
    /// <param name="x">x坐标值</param>
    /// <param name="y">y坐标值</param>
    /// <returns>指定坐标点处的单元格区域对象，如果坐标点不在工作表区域内则返回null</returns>
    object? RangeFromPoint(int x, int y);

    /// <summary>
    /// 获取选中的工作表集合
    /// </summary>
    IExcelSheets? SelectedSheets { get; }

    /// <summary>
    /// 获取关联的工作表
    /// </summary>
    IExcelWorksheet? ActiveSheet { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取父工作簿<see cref="IExcelWorkbook"/>对象
    /// </summary>
    IExcelWorkbook? ParentWorkbook { get; }

    /// <summary>
    /// 获取窗口索引
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取是否为活动窗口
    /// </summary>
    bool IsActive { get; }

    /// <summary>
    /// 选择指定范围
    /// </summary>
    /// <param name="rangeAddress">范围地址，如 "A1:B10"</param>
    void SelectRange(string rangeAddress);

    /// <summary>
    /// 滚动到指定范围
    /// </summary>
    /// <param name="rangeAddress">范围地址</param>
    void ScrollToRange(string rangeAddress);

    /// <summary>
    /// 刷新窗口显示
    /// </summary>
    void Refresh();

    /// <summary>
    /// 保存窗口布局
    /// </summary>
    void SaveLayout();

    /// <summary>
    /// 恢复窗口布局
    /// </summary>
    void RestoreLayout();

    /// <summary>
    /// 创建当前窗口的新实例
    /// </summary>
    /// <returns>新窗口对象</returns>
    IExcelWindow NewWindow();

    /// <summary>
    /// 大范围滚动窗口内容
    /// </summary>
    /// <param name="down">向下滚动页数</param>
    /// <param name="up">向上滚动页数</param>
    /// <param name="right">向右滚动页数</param>
    /// <param name="left">向左滚动页数</param>
    void LargeScroll(int? down = 0, int? up = 0, int? right = 0, int? left = 0);

    /// <summary>
    /// 小范围滚动窗口内容
    /// </summary>
    /// <param name="down">向下滚动行数</param>
    /// <param name="up">向上滚动行数</param>
    /// <param name="right">向右滚动列数</param>
    /// <param name="left">向左滚动列数</param>
    void SmallScroll(int? down = 0, int? up = 0, int? right = 0, int? left = 0);

    /// <summary>
    /// 将水平坐标点转换为屏幕像素值
    /// </summary>
    /// <param name="points">点坐标</param>
    /// <returns>像素值</returns>
    int PointsToScreenPixelsX(int points);

    /// <summary>
    /// 将垂直坐标点转换为屏幕像素值
    /// </summary>
    /// <param name="points">点坐标</param>
    /// <returns>像素值</returns>
    int PointsToScreenPixelsY(int points);

    /// <summary>
    /// 执行打印预览
    /// </summary>
    void PrintPreview();

    /// <summary>
    /// 打印窗口内容
    /// </summary>
    /// <param name="preview">是否预览</param>
    void PrintOut(bool preview = false);

    /// <summary>
    /// 封装Excel窗口打印方法
    /// </summary>
    /// <param name="copies">打印份数</param>
    /// <param name="preview">是否预览</param>
    /// <param name="activePrinter">打印机名称</param>
    /// <param name="printToFile">是否打印到文件</param>
    /// <param name="collate">是否逐份打印</param>
    void PrintOut(
        int copies = 1,
        bool preview = false,
        string? activePrinter = null,
        bool printToFile = false,
        bool collate = true);
}