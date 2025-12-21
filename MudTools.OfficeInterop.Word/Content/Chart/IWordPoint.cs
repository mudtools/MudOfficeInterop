//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中图表的一个数据点接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordPoint : IDisposable
{

    /// <summary>
    /// 获取与指定对象相关联的应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }


    /// <summary>
    /// 获取该对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 返回指定数据点的边框格式
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 返回指定数据点的数据标签
    /// </summary>
    IWordDataLabel? DataLabel { get; }

    /// <summary>
    /// 返回指定数据点的填充格式
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 返回指定对象的内部属性
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取指定对象的格式属性
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 返回或设置饼图数据点从饼图中心分离的距离（以磅为单位）
    /// </summary>
    int Explosion { get; set; }

    /// <summary>
    /// 返回或设置一个布尔值，该值决定是否显示数据标签
    /// </summary>
    bool HasDataLabel { get; set; }

    /// <summary>
    /// 如果指定数据点的负数数值使用相反的颜色，则为 true
    /// </summary>
    bool InvertIfNegative { get; set; }

    /// <summary>
    /// 返回或设置图表图片单位大小（以磅为单位）
    /// </summary>
    double PictureUnit2 { get; set; }

    /// <summary>
    /// 获取数据点的高度（以磅为单位）
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取数据点的宽度（以磅为单位）
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取数据点左边位置坐标（以磅为单位）
    /// </summary>
    double Left { get; }

    /// <summary>
    /// 获取数据点顶部位置坐标（以磅为单位）
    /// </summary>
    double Top { get; }

    /// <summary>
    /// 获取对象的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 返回或设置数据标记的背景颜色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerBackgroundColor { get; set; }

    /// <summary>
    /// 返回或设置数据标记的前景颜色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerForegroundColor { get; set; }

    /// <summary>
    /// 返回或设置数据标记的背景色索引
    /// </summary>
    XlColorIndex MarkerBackgroundColorIndex { get; set; }

    /// <summary>
    /// 返回或设置数据标记的前景色索引
    /// </summary>
    XlColorIndex MarkerForegroundColorIndex { get; set; }

    /// <summary>
    /// 返回或设置数据标记的样式
    /// </summary>
    XlMarkerStyle MarkerStyle { get; set; }

    /// <summary>
    /// 返回或设置图片类型
    /// </summary>
    XlChartPictureType PictureType { get; set; }

    /// <summary>
    /// 返回或设置数据标记的大小
    /// </summary>
    int MarkerSize { get; set; }

    /// <summary>
    /// 如果数据点在次要坐标轴上绘制，则为 true
    /// </summary>
    bool SecondaryPlot { get; set; }

    /// <summary>
    /// 如果图片应用于三维条形图或柱形图的正面，则为 true
    /// </summary>
    bool ApplyPictToFront { get; set; }

    /// <summary>
    /// 如果图片应用于三维条形图或柱形图的一端，则为 true
    /// </summary>
    bool ApplyPictToEnd { get; set; }

    /// <summary>
    /// 返回或设置一个布尔值，表示数据点是否有阴影效果
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 如果图片应用于三维条形图或柱形图的侧面，则为 true
    /// </summary>
    bool ApplyPictToSides { get; set; }

    /// <summary>
    /// 返回或设置一个布尔值，表示数据点是否有三维效果
    /// </summary>
    bool Has3DEffect { get; set; }

    /// <summary>
    /// 返回或设置图表图片单位大小（旧版本兼容属性）
    /// </summary>
    double PictureUnit { get; set; }

    /// <summary>
    /// 向指定数据点应用数据标签
    /// </summary>
    /// <param name="type">数据标签的类型</param>
    /// <param name="legendKey">如果为 true，则在数据标签旁边显示图例项标示</param>
    /// <param name="autoText">如果为 true，则自动生成对象的文本内容</param>
    /// <param name="hasLeaderLines">对于图表，如果为 true 则数据标签具有引导线</param>
    /// <param name="showSeriesName">如果为 true，则在数据标签中显示系列名称</param>
    /// <param name="showCategoryName">如果为 true，则在数据标签中显示分类名称</param>
    /// <param name="showValue">如果为 true，则在数据标签中显示值</param>
    /// <param name="showPercentage">如果为 true，则在数据标签中显示百分比</param>
    /// <param name="showBubbleSize">如果为 true，则在数据标签中显示气泡大小</param>
    /// <param name="separator">数据标签中使用的分隔符</param>
    /// <returns>返回操作结果对象</returns>
    object? ApplyDataLabels(XlDataLabelsType type = XlDataLabelsType.xlDataLabelsShowValue,
                        bool? legendKey = null, string? autoText = null, bool? hasLeaderLines = null,
                        bool? showSeriesName = null, bool? showCategoryName = null,
                        bool? showValue = null, bool? showPercentage = null,
                        bool? showBubbleSize = null, string? separator = null);
                        
    /// <summary>
    /// 返回饼图数据点的位置坐标
    /// </summary>
    /// <param name="loc">指定要获取的位置类型</param>
    /// <param name="Index">指定要获取的点索引</param>
    /// <returns>返回指定位置的坐标值</returns>
    double? PieSliceLocation(XlPieSliceLocation loc, XlPieSliceIndex Index = XlPieSliceIndex.xlOuterCenterPoint);

    /// <summary>
    /// 清除指定对象的格式设置
    /// </summary>
    /// <returns>返回操作结果对象</returns>
    object? ClearFormats();

    /// <summary>
    /// 复制对象到剪贴板
    /// </summary>
    /// <returns>返回操作结果对象</returns>
    object? Copy();

    /// <summary>
    /// 从剪贴板粘贴对象
    /// </summary>
    /// <returns>返回操作结果对象</returns>
    object? Paste();

    /// <summary>
    /// 删除对象
    /// </summary>
    /// <returns>返回操作结果对象</returns>
    object? Delete();

    /// <summary>
    /// 选择对象
    /// </summary>
    /// <returns>返回操作结果对象</returns>
    object? Select();
}