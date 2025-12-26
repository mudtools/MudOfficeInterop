//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Series 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Series 的安全访问和操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSeries : IOfficeObject<IExcelSeries>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置数据系列的名称
    /// 对应 Series.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取数据系列的父对象 (通常是 SeriesCollection)
    /// 对应 Series.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据系列所在的 Application 对象
    /// 对应 Series.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置系列的图表类型
    /// 对应 Series.ChartType 属性
    /// </summary>
    XlChartType ChartType { get; set; }

    /// <summary>
    /// 获取或设置数据系列的坐标轴组
    /// 对应 Series.AxisGroup 属性
    /// </summary>
    XlAxisGroup AxisGroup { get; set; }

    /// <summary>
    /// 获取或设置系列的公式
    /// 对应 Series.Formula 属性
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置系列的本地化公式
    /// 对应 Series.FormulaLocal 属性
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置系列的R1C1引用样式公式
    /// 对应 Series.FormulaR1C1 属性
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置系列的本地化R1C1引用样式公式
    /// 对应 Series.FormulaR1C1Local 属性
    /// </summary>
    string FormulaR1C1Local { get; set; }
    #endregion

    #region 数据属性
    /// <summary>
    /// 获取或设置系列的X轴值区域
    /// 对应 Series.XValues 属性，可以是 object[] 或 IExcelRange
    /// </summary>
    object? XValues { get; set; }

    /// <summary>
    /// 获取或设置系列的Y轴值区域
    /// 对应 Series.Values 属性，可以是 object[] 或 IExcelRange
    /// </summary>
    object? Values { get; set; }

    /// <summary>
    /// 获取或设置系列的气泡大小值区域 (气泡图)， 可以是 object[] 或 IExcelRange
    /// 对应 Series.BubbleSizes 属性
    /// </summary>
    object? BubbleSizes { get; set; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取绘图区的字体对象
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取系列的背景填充对象
    /// 对应 Series.Format.Fill 属性
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取系列的边框对象
    /// 对应 Series.Format.Line 属性
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取或设置系列的标记样式
    /// 对应 Series.MarkerStyle 属性
    /// </summary>
    XlMarkerStyle MarkerStyle { get; set; }

    /// <summary>
    /// 获取或设置系列的标记大小
    /// 对应 Series.MarkerSize 属性
    /// </summary>
    int MarkerSize { get; set; }

    /// <summary>
    /// 获取或设置系列的标记背景色
    /// 对应 Series.MarkerBackgroundColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerBackgroundColor { get; set; }

    /// <summary>
    /// 获取或设置系列的标记背景色索引
    /// 对应 Series.MarkerBackgroundColorIndex 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColorIndex MarkerBackgroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置系列的标记前景色
    /// 对应 Series.MarkerForegroundColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerForegroundColor { get; set; }

    /// <summary>
    /// 获取或设置系列的标记前景色索引
    /// 对应 Series.MarkerForegroundColorIndex 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColorIndex MarkerForegroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置图片图表类型
    /// 对应 Series.PictureType 属性
    /// </summary>
    XlChartPictureType PictureType { get; set; }

    /// <summary>
    /// 获取或设置图片单位大小（整数形式）
    /// 对应 Series.PictureUnit 属性
    /// </summary>
    int PictureUnit { get; set; }

    /// <summary>
    /// 获取或设置数据系列的类型
    /// 对应 Series.Type 属性
    /// </summary>
    int Type { get; set; }

    /// <summary>
    /// 获取或设置三维条形图形状
    /// 对应 Series.BarShape 属性
    /// </summary>
    XlBarShape BarShape { get; set; }

    /// <summary>
    /// 获取或设置是否将图片应用于柱体的侧面
    /// 对应 Series.ApplyPictToSides 属性
    /// </summary>
    bool ApplyPictToSides { get; set; }

    /// <summary>
    /// 获取或设置是否将图片应用于柱体的正面
    /// 对应 Series.ApplyPictToFront 属性
    /// </summary>
    bool ApplyPictToFront { get; set; }

    /// <summary>
    /// 获取或设置是否将图片应用于柱体的末端
    /// 对应 Series.ApplyPictToEnd 属性
    /// </summary>
    bool ApplyPictToEnd { get; set; }

    /// <summary>
    /// 获取或设置数据系列是否具有三维效果
    /// 对应 Series.Has3DEffect 属性
    /// </summary>
    bool Has3DEffect { get; set; }

    /// <summary>
    /// 获取或设置数据系列是否具有阴影效果
    /// 对应 Series.Shadow 属性
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置图片单位大小（浮点数形式）
    /// 对应 Series.PictureUnit2 属性
    /// </summary>
    double PictureUnit2 { get; set; }

    /// <summary>
    /// 获取绘图区的颜色索引
    /// 对应 Series.PlotColorIndex 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColorIndex PlotColorIndex { get; }

    /// <summary>
    /// 获取或设置反转颜色
    /// 对应 Series.InvertColor 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color InvertColor { get; set; }

    /// <summary>
    /// 获取或设置反转颜色索引
    /// 对应 Series.InvertColorIndex 属性
    /// </summary>
    int InvertColorIndex { get; set; }

    /// <summary>
    /// 获取数据标签的引导线对象
    /// 对应 Series.LeaderLines 属性
    /// </summary>
    IExcelLeaderLines LeaderLines { get; }

    /// <summary>
    /// 获取或设置数据系列是否被筛选隐藏
    /// 对应 Series.IsFiltered 属性
    /// </summary>
    bool IsFiltered { get; set; }


    /// <summary>
    /// 获取或设置系列的平滑线
    /// 对应 Series.Smooth 属性
    /// </summary>
    bool Smooth { get; set; }

    /// <summary>
    /// 获取或设置系列的绘制顺序
    /// 对应 Series.PlotOrder 属性
    /// </summary>
    int PlotOrder { get; set; }

    /// <summary>
    /// 获取或设置数据系列的分离程度（饼图扇区分离距离）
    /// 对应 Series.Explosion 属性
    /// </summary>
    int Explosion { get; set; }

    /// <summary>
    /// 获取或设置当值为负数时是否反转颜色
    /// 对应 Series.InvertIfNegative 属性
    /// </summary>
    bool InvertIfNegative { get; set; }

    #endregion

    #region 状态属性
    /// <summary>
    /// 获取或设置是否包含在图例中
    /// 对应 Series.HasLeaderLines 属性 (注意：原文可能意为 HasLegendKey)
    /// 更准确的是通过 DataLabels.ShowLegendKey 或 Chart.Legend 来控制
    /// 这里保留此属性名，但实现需注意其实际含义可能需要通过其他方式控制
    /// </summary>
    bool HasLeaderLines { get; set; }

    /// <summary>
    /// 获取或设置是否显示数据标签
    /// 对应 Series.HasDataLabels 属性
    /// </summary>
    bool HasDataLabels { get; set; }

    /// <summary>
    /// 获取或设置是否显示错误线
    /// 对应 Series.HasErrorBars 属性
    /// </summary>
    bool HasErrorBars { get; set; }
    #endregion

    #region 图表元素 (子对象)

    /// <summary>
    /// 获取样式的内部格式对象
    /// 对应 Style.Interior 属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取系列的X轴误差线
    /// 对应 Series.ErrorBars 属性 (通常指Y轴误差线，X轴需要特殊获取)
    /// </summary>
    IExcelErrorBars? ErrorBars { get; }

    #endregion

    #region 操作方法
    /// <summary>
    /// 为数据系列添加误差线
    /// 对应 Series.ErrorBar 方法
    /// </summary>
    /// <param name="direction">误差线方向，水平(X轴)或垂直(Y轴)</param>
    /// <param name="include">误差线包含范围，正误差线、负误差线或两者都包含</param>
    /// <param name="type">误差线类型，如固定值、百分比、标准偏差等</param>
    /// <param name="amount">正误差线的值，根据类型不同含义不同，如为null则使用默认值</param>
    /// <param name="minusValues">负误差线的值，根据类型不同含义不同，如为null则使用默认值</param>
    [ReturnValueConvert]
    IExcelErrorBars? ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type, object? amount = null, object? minusValues = null);

    /// <summary>
    /// 获取数据系列的所有趋势线集合
    /// 对应 Series.Trendlines() 方法
    /// </summary>
    /// <returns>趋势线集合 <see cref="IExcelTrendlines"/> 对象，如果获取失败则返回 null</returns>
    [ReturnValueConvert]
    IExcelTrendlines? Trendlines();

    /// <summary>
    /// 获取指定类型的趋势线对象
    /// 对应 Series.Trendlines(type) 方法
    /// </summary>
    /// <param name="trendlineType">趋势线类型 <see cref="XlTrendlineType"/></param>
    /// <returns>趋势线对象 <see cref="IExcelTrendline"/>，如果获取失败则返回 null</returns>
    [ReturnValueConvert]
    IExcelTrendline? Trendlines(XlTrendlineType trendlineType);

    /// <summary>
    /// 获取数据系列的所有数据标签集合
    /// 对应 Series.DataLabels() 方法
    /// </summary>
    /// <returns>数据标签集合 <see cref="IExcelDataLabels"/> 对象，如果获取失败则返回 null</returns>
    [ReturnValueConvert]
    IExcelDataLabels? DataLabels();

    /// <summary>
    /// 获取数据系列中特定索引的数据标签
    /// 对应 Series.DataLabels(object) 方法
    /// </summary>
    /// <param name="index">指定要返回的数据标签的索引号或标识</param>
    /// <returns>指定的数据标签 <see cref="IExcelDataLabel"/> 对象，如果获取失败则返回 null</returns>
    [ReturnValueConvert]
    IExcelDataLabel? DataLabels(int index);

    /// <summary>
    /// 获取数据系列中特定索引的数据标签
    /// 对应 Series.DataLabels(object) 方法
    /// </summary>
    /// <param name="name">指定要返回的数据标签的索引号或标识</param>
    /// <returns>指定的数据标签 <see cref="IExcelDataLabel"/> 对象，如果获取失败则返回 null</returns>
    [ReturnValueConvert]
    IExcelDataLabel? DataLabels(string name);

    /// <summary>
    /// 选择数据系列
    /// 对应 Series.Select 方法
    /// </summary>
    object? Select();

    /// <summary>
    /// 删除数据系列
    /// 对应 Series.Delete 方法
    /// </summary>
    object? Delete();

    /// <summary>
    /// 清除数据系列格式
    /// 对应 Series.ClearFormats 方法
    /// </summary>
    object? ClearFormats();

    /// <summary>
    /// 复制数据系列
    /// 对应 Series.Copy 方法
    /// </summary>
    object? Copy();

    /// <summary>
    /// 获取数据系列中的所有数据点
    /// </summary>
    /// <returns>IExcelPoints 对象，包含系列中的所有数据点</returns>
    [ReturnValueConvert]
    IExcelPoints? Points();

    /// <summary>
    /// 根据名称获取数据系列中的特定数据点
    /// </summary>
    /// <param name="name">数据点的名称</param>
    /// <returns>IExcelPoint 对象，表示指定名称的数据点</returns>
    [ReturnValueConvert]
    IExcelPoint? Points(string name);

    /// <summary>
    /// 根据索引获取数据系列中的特定数据点
    /// </summary>
    /// <param name="index">数据点在系列中的索引位置</param>
    /// <returns>IExcelPoint 对象，表示指定索引位置的数据点</returns>
    [ReturnValueConvert]
    IExcelPoint? Points(int index);
    #endregion

    #region 图表操作
    /// <summary>
    /// 应用数据标签
    /// 对应 Series.ApplyDataLabels 方法
    /// </summary>
    /// <param name="type">标签类型</param>
    /// <param name="legendKey">是否显示图例项标示</param>
    /// <param name="autoText">是否自动生成文本</param>
    /// <param name="hasLeaderLines">是否显示引导线</param>
    /// <param name="showSeriesName">是否显示系列名称</param>
    /// <param name="showCategoryName">是否显示分类名称</param>
    /// <param name="showValue">是否显示值</param>
    /// <param name="showPercentage">是否显示百分比</param>
    /// <param name="showBubbleSize">是否显示气泡大小</param>
    /// <param name="separator">分隔符</param>
    void ApplyDataLabels(XlDataLabelsType type = XlDataLabelsType.xlDataLabelsShowValue,
                                  bool? legendKey = null, bool? autoText = null,
                                  bool? hasLeaderLines = null, bool? showSeriesName = null,
                                  bool? showCategoryName = null, bool? showValue = null,
                                  bool? showPercentage = null, bool? showBubbleSize = null,
                                  string? separator = null);


    /// <summary>
    /// 将自定义图表类型应用于当前数据系列
    /// </summary>
    /// <param name="ChartType">指定要应用的图表类型，该类型定义了数据系列的可视化表现形式</param>
    void ApplyCustomType(XlChartType ChartType);

    #endregion
}
