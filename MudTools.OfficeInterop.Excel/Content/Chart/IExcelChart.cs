//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// Excel Chart 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Chart 的安全访问和操作
/// </summary>
public interface IExcelChart : IExcelComSheet, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取或设置图表的类型
    /// 对应 Chart.ChartType 属性
    /// </summary>
    MsoChartType ChartType { get; set; }

    /// <summary>
    /// 获取图表的代码名称
    /// 对应 Chart.CodeName 属性
    /// </summary>
    string CodeName { get; }
    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置图表的旋转角度
    /// 对应 Chart.Rotation 属性
    /// </summary>
    double Rotation { get; set; }

    /// <summary>
    /// 获取或设置三维图表的视角高度
    /// 对应 Chart.Elevation 属性，以度为单位，范围通常在-90到90之间
    /// </summary>
    int Elevation { get; set; }

    /// <summary>
    /// 获取或设置三维图表的深度百分比
    /// 对应 Chart.DepthPercent 属性，表示图表深度相对于其宽度的百分比，范围通常在5到500之间
    /// </summary>
    int DepthPercent { get; set; }

    /// <summary>
    /// 获取或设置三维图表中数据系列之间的距离
    /// 对应 Chart.GapDepth 属性，以图表宽度的百分比表示，范围通常在0到500之间
    /// </summary>
    int GapDepth { get; set; }

    /// <summary>
    /// 获取或设置三维图表的透视角度
    /// 对应 Chart.Perspective 属性，以度为单位，范围通常在0到100之间
    /// </summary>
    int Perspective { get; set; }

    #endregion

    #region 数据源
    /// <summary>
    /// 获取或设置绘制方式
    /// 对应 Chart.PlotBy 属性
    /// </summary>
    XlRowCol PlotBy { get; set; }

    /// <summary>
    /// 获取或设置是否包含标题行
    /// 对应 Chart.HasTitle 属性
    /// </summary>
    bool HasTitle { get; set; }

    /// <summary>
    /// 获取或设置图表标题
    /// 对应 Chart.ChartTitle 属性
    /// </summary>
    IExcelChartTitle? ChartTitle { get; }

    /// <summary>
    /// 获取或设置是否包含图例
    /// 对应 Chart.HasLegend 属性
    /// </summary>
    bool HasLegend { get; set; }

    /// <summary>
    /// 获取或设置图例位置
    /// 对应 Chart.Legend.Position 属性
    /// </summary>
    XlLegendPosition LegendPosition { get; set; }

    #endregion

    #region 图表元素
    /// <summary>
    /// 获取图表的绘图区对象
    /// 对应 Chart.PlotArea 属性
    /// </summary>
    IExcelPlotArea? PlotArea { get; }

    /// <summary>
    /// 获取图表的图表区对象
    /// 对应 Chart.ChartArea 属性
    /// </summary>
    IExcelChartArea? ChartArea { get; }

    /// <summary>
    /// 获取图表中所有三维折线图组的第一个图表组
    /// 对应 Chart.Line3DGroup 属性
    /// </summary>
    IExcelChartGroup? Line3DGroup { get; }

    /// <summary>
    /// 为图表中的所有数据系列应用数据标签
    /// </summary>
    /// <param name="type">指定要显示的数据标签类型，默认为显示数值</param>
    /// <param name="legendKey">图例标识，可选参数</param>
    /// <param name="autoText">自动文本，可选参数</param>
    /// <param name="hasLeaderLines">是否具有引导线，可选参数</param>
    /// <param name="showSeriesName">是否显示系列名称，可选参数</param>
    /// <param name="showCategoryName">是否显示分类名称，可选参数</param>
    /// <param name="showValue">是否显示值，可选参数</param>
    /// <param name="showPercentage">是否显示百分比，可选参数</param>
    /// <param name="showBubbleSize">是否显示气泡大小，可选参数</param>
    /// <param name="separator">分隔符，可选参数</param>
    void ApplyDataLabels(XlDataLabelsType type = XlDataLabelsType.xlDataLabelsShowValue,
                               bool? legendKey = null, string? autoText = null, bool? hasLeaderLines = null, bool? showSeriesName = null,
                               bool? showCategoryName = null, bool? showValue = null, bool? showPercentage = null,
                               bool? showBubbleSize = null, string? separator = null);
    /// <summary>
    /// 返回图表中所有图表组的集合或指定索引的图表组
    /// </summary>
    /// <param name="Index">图表组的索引号，从1开始，如果省略则返回所有图表组</param>
    /// <returns>图表组对象或图表组集合</returns>
    object? ChartGroups(object? Index = null);

    /// <summary>
    /// 返回图表中所有柱形图组的集合或指定索引的柱形图组
    /// </summary>
    /// <param name="Index">柱形图组的索引号，从1开始，如果省略则返回所有柱形图组</param>
    /// <returns>柱形图组对象或柱形图组集合</returns>
    object? BarGroups(object? Index = null);

    /// <summary>
    /// 返回图表中所有折线图组的集合或指定索引的折线图组
    /// </summary>
    /// <param name="Index">折线图组的索引号，从1开始，如果省略则返回所有折线图组</param>
    /// <returns>折线图组对象或折线图组集合</returns>
    object? LineGroups(object? Index = null);

    /// <summary>
    /// 获取图表的坐标轴集合
    /// 对应 Chart.Axes 函数
    /// </summary>
    IExcelAxis? Axes(XlAxisType? axisType, XlAxisGroup axisGroup = XlAxisGroup.xlPrimary);

    /// <summary>
    /// 获取图表的坐标轴集合
    /// 对应 Chart.Axes 函数
    /// </summary>
    IExcelAxes? Axes();


    /// <summary>
    /// 获取图表的图表标题对象
    /// 对应 Chart.ChartTitle 属性
    /// </summary>
    IExcelChartTitle? ChartTitleObject { get; }

    /// <summary>
    /// 获取图表的图例对象
    /// 对应 Chart.Legend 属性
    /// </summary>
    IExcelLegend? Legend { get; }

    /// <summary>
    /// 获取图表的数据标签集合
    /// 对应 Chart.DataTable 属性
    /// </summary>
    IExcelDataTable? DataTable { get; }
    #endregion

    #region 图表设置

    /// <summary>
    /// 获取或设置是否显示数据表
    /// 对应 Chart.HasDataTable 属性
    /// </summary>
    bool HasDataTable { get; set; }

    /// <summary>
    /// 获取或设置图表样式
    /// 对应 Chart.ChartStyle 属性
    /// </summary>
    MsoChartType ChartStyle { get; set; }

    #endregion

    #region 操作方法   
    /// <summary>
    /// 旋转图表
    /// </summary>
    /// <param name="angle">旋转角度</param>
    void Rotate(double angle);
    #endregion

    #region 图表操作

    /// <summary>
    /// 设置图表的背景图片
    /// </summary>
    /// <param name="filename">图片文件的路径和文件名</param>
    void SetBackgroundPicture(string filename);

    /// <summary>
    /// 设置图表数据源
    /// </summary>
    /// <param name="sourceData">数据源区域</param>
    /// <param name="plotBy">绘制方式</param>
    void SetSourceData(IExcelRange sourceData, XlRowCol plotBy = XlRowCol.xlRows);


    /// <summary>
    /// 应用图表布局
    /// </summary>
    /// <param name="layout">布局编号</param>
    void ApplyLayout(int layout);

    /// <summary>
    /// 刷新图表
    /// </summary>
    void Refresh();

    /// <summary>
    /// 清除图表格式
    /// </summary>
    void ClearFormats();

    #endregion

    #region 格式设置

    /// <summary>
    /// 获取图表的数据系列集合
    /// </summary>
    IExcelSeriesCollection? SeriesCollection();

    /// <summary>
    /// 获取图表的数据系列集合
    /// </summary>
    IExcelSeries? SeriesCollection(int index);

    /// <summary>
    /// 设置图表标题
    /// </summary>
    /// <param name="title">标题文本</param>
    void SetTitle(string title);

    /// <summary>
    /// 设置图例位置
    /// </summary>
    /// <param name="position">图例位置</param>
    void SetLegendPosition(XlLegendPosition position);

    /// <summary>
    /// 设置数据标签
    /// </summary>
    /// <param name="show">是否显示</param>
    void SetDataLabels(bool show);

    /// <summary>
    /// 设置图表背景色
    /// </summary>
    /// <param name="color">背景色</param>
    void SetBackgroundColor(int color);

    /// <summary>
    /// 设置图表前景色
    /// </summary>
    /// <param name="color">前景色</param>
    void SetForegroundColor(int color);

    #endregion

    #region 导出和转换

    /// <summary>
    /// 导出图表到图片文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="format">图片格式</param>
    /// <param name="overwrite">是否覆盖已存在文件</param>
    /// <returns>是否导出成功</returns>
    bool ExportToImage(string filename, string format = "png", bool overwrite = true);

    /// <summary>
    /// 获取图表的图片字节数据
    /// </summary>
    /// <param name="format">图片格式</param>
    /// <returns>图片字节数组</returns>
    byte[] GetImageBytes(string format = "png");

    #endregion  


    #region 公共事件

    /// <summary>
    /// 图表被激活时触发
    /// </summary>
    event ChartActivateEventHandler ChartActivate;
    /// <summary>
    /// 图表失活时触发
    /// </summary>
    event ChartDeactivateEventHandler Deactivate;

    /// <summary>
    /// 用户在图表上选择任意元素时触发（如点击数据点、图例、标题等）
    /// </summary>
    event ChartSelectEventHandler ChartSelect;

    /// <summary>
    /// 用户双击图表前触发，设置 Cancel = true 可阻止默认行为（如打开“设置数据系列”对话框）
    /// </summary>
    event ChartBeforeDoubleClickEventHandler BeforeDoubleClick;

    /// <summary>
    /// 用户右键单击图表前触发，设置 Cancel = true 可阻止弹出上下文菜单
    /// </summary>
    event ChartBeforeRightClickEventHandler BeforeRightClick;

    /// <summary>
    /// 当图表中的数据系列发生变化时触发（例如单元格数据更新导致图表重绘）
    /// </summary>
    event ChartSeriesChangeEventHandler SeriesChange;
    #endregion
}
