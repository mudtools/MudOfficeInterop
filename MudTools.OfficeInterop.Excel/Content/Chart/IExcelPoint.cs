//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 图表中的数据点接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelPoint : IDisposable
{
    /// <summary>
    /// 获取对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取包含该对象的应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取数据点的边框属性
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取数据点的内部填充属性
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取数据点的填充格式属性
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取数据点的图表格式属性
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取或设置数据点从饼图中心移出的距离
    /// </summary>
    int Explosion { get; set; }

    /// <summary>
    /// 获取或设置数据点是否具有数据标签
    /// </summary>
    bool HasDataLabel { get; set; }

    /// <summary>
    /// 获取或设置负值数据点是否使用相反颜色显示
    /// </summary>
    bool InvertIfNegative { get; set; }

    /// <summary>
    /// 获取或设置标记背景色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerBackgroundColor { get; set; }

    /// <summary>
    /// 获取或设置标记背景色索引
    /// </summary>
    XlColorIndex MarkerBackgroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置标记前景色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerForegroundColor { get; set; }

    /// <summary>
    /// 获取或设置标记前景色索引
    /// </summary>
    XlColorIndex MarkerForegroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置标记大小
    /// </summary>
    int MarkerSize { get; set; }

    /// <summary>
    /// 获取或设置图片类型
    /// </summary>
    XlChartPictureType PictureType { get; set; }

    /// <summary>
    /// 获取或设置标记样式
    /// </summary>
    XlMarkerStyle MarkerStyle { get; set; }

    /// <summary>
    /// 获取或设置图片单位
    /// </summary>
    int PictureUnit { get; set; }

    /// <summary>
    /// 获取或设置是否将图片应用到侧面
    /// </summary>
    bool ApplyPictToSides { get; set; }

    /// <summary>
    /// 获取或设置是否将图片应用到正面
    /// </summary>
    bool ApplyPictToFront { get; set; }

    /// <summary>
    /// 获取或设置是否将图片应用到末端
    /// </summary>
    bool ApplyPictToEnd { get; set; }

    /// <summary>
    /// 获取或设置数据点是否有阴影效果
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置数据点是否在辅助坐标轴上显示
    /// </summary>
    bool SecondaryPlot { get; set; }

    /// <summary>
    /// 获取或设置数据点是否具有三维效果
    /// </summary>
    bool Has3DEffect { get; set; }

    /// <summary>
    /// 获取或设置图片单位值（双精度浮点型）
    /// </summary>
    double PictureUnit2 { get; set; }

    /// <summary>
    /// 获取数据点的高度
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取数据点的宽度
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取数据点的顶部位置
    /// </summary>
    double Top { get; }

    /// <summary>
    /// 获取数据点的左侧位置
    /// </summary>
    double Left { get; }

    /// <summary>
    /// 获取数据点的名称
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 清除数据点的格式
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? ClearFormats();

    /// <summary>
    /// 复制数据点
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? Copy();

    /// <summary>
    /// 删除数据点
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? Delete();

    /// <summary>
    /// 粘贴内容到数据点
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? Paste();

    /// <summary>
    /// 选择数据点
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? Select();

    /// <summary>
    /// 应用数据标签到数据点
    /// </summary>
    /// <param name="type">数据标签类型</param>
    /// <param name="legendKey">是否显示图例项标示</param>
    /// <param name="autoText">是否自动生成文本</param>
    /// <param name="hasLeaderLines">是否具有引导线</param>
    /// <param name="showSeriesName">是否显示系列名称</param>
    /// <param name="showCategoryName">是否显示分类名称</param>
    /// <param name="showValue">是否显示值</param>
    /// <param name="showPercentage">是否显示百分比</param>
    /// <param name="showBubbleSize">是否显示气泡大小</param>
    /// <param name="separator">分隔符</param>
    /// <returns>操作结果对象</returns>
    object? ApplyDataLabels(XlDataLabelsType type = XlDataLabelsType.xlDataLabelsShowValue,
         bool? legendKey = null, bool? autoText = null,
         bool? hasLeaderLines = null, bool? showSeriesName = null,
         bool? showCategoryName = null, bool? showValue = null,
         bool? showPercentage = null, bool? showBubbleSize = null,
         string? separator = null);

    /// <summary>
    /// 获取饼图切片的位置信息
    /// </summary>
    /// <param name="loc">饼图切片位置类型</param>
    /// <param name="Index">饼图切片索引</param>
    /// <returns>位置坐标的双精度浮点值</returns>
    double? PieSliceLocation(XlPieSliceLocation loc, XlPieSliceIndex Index = XlPieSliceIndex.xlOuterCenterPoint);
}