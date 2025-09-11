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
public interface IExcelChart : ICommonWorksheet, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取或设置图表的类型
    /// 对应 Chart.ChartType 属性
    /// </summary>
    MsoChartType ChartType { get; set; }

    /// <summary>
    /// 获取或设置工作表是否可见
    /// 对应 Worksheet.Visible 属性
    /// </summary>
    XlSheetVisibility Visible { get; set; }

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

    #endregion

    #region 数据源
    /// <summary>
    /// 获取或设置绘制方式
    /// 对应 Chart.PlotBy 属性
    /// </summary>
    int PlotBy { get; set; }

    /// <summary>
    /// 获取或设置是否包含标题行
    /// 对应 Chart.HasTitle 属性
    /// </summary>
    bool HasTitle { get; set; }

    /// <summary>
    /// 获取或设置图表标题
    /// 对应 Chart.ChartTitle 属性
    /// </summary>
    string ChartTitle { get; set; }

    /// <summary>
    /// 获取或设置是否包含图例
    /// 对应 Chart.HasLegend 属性
    /// </summary>
    bool HasLegend { get; set; }

    /// <summary>
    /// 获取或设置图例位置
    /// 对应 Chart.Legend.Position 属性
    /// </summary>
    int LegendPosition { get; set; }

    #endregion

    #region 图表元素

    IExcelShapes Shapes { get; }

    /// <summary>
    /// 获取图表的绘图区对象
    /// 对应 Chart.PlotArea 属性
    /// </summary>
    IExcelPlotArea PlotArea { get; }

    /// <summary>
    /// 获取图表的图表区对象
    /// 对应 Chart.ChartArea 属性
    /// </summary>
    IExcelChartArea ChartArea { get; }

    /// <summary>
    /// 获取图表的坐标轴集合
    /// 对应 Chart.Axes 属性
    /// </summary>
    IExcelAxes Axes { get; }

    /// <summary>
    /// 获取图表的图表标题对象
    /// 对应 Chart.ChartTitle 属性
    /// </summary>
    IExcelChartTitle ChartTitleObject { get; }

    /// <summary>
    /// 获取图表的图例对象
    /// 对应 Chart.Legend 属性
    /// </summary>
    IExcelLegend Legend { get; }

    /// <summary>
    /// 获取图表的数据标签集合
    /// 对应 Chart.DataTable 属性
    /// </summary>
    IExcelDataTable DataTable { get; }
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
    int ChartStyle { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 将工作表另存为xlsx文件。
    /// </summary>
    /// <param name="filePath"></param>
    void SaveAs(string filePath);
    /// <summary>
    /// 旋转图表
    /// </summary>
    /// <param name="angle">旋转角度</param>
    void Rotate(double angle);
    #endregion

    #region 图表操作

    /// <summary>
    /// 设置图表数据源
    /// </summary>
    /// <param name="sourceData">数据源区域</param>
    /// <param name="plotBy">绘制方式</param>
    void SetSourceData(IExcelRange sourceData, int plotBy = 1);


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
    /// 清除图表内容
    /// </summary>
    void Clear();

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

    #region 高级功能

    /// <summary>
    /// 打印图表
    /// </summary>
    /// <param name="preview">是否打印预览</param>
    void PrintOut(bool preview = false);
    #endregion
}
