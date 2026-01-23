//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


using System.Drawing;

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中图表的图例键接口
/// </summary>
/// <remarks>
/// 该接口封装了 Microsoft Word 图表中图例键的相关属性和方法，
/// 提供对图例键格式、样式和相关对象的访问能力。
/// 图例键是图表中用来标识不同数据系列的视觉标记。
/// </remarks>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLegendKey : IOfficeObject<IWordLegendKey>, IDisposable
{

    /// <summary>
    /// 获取与该图例键关联的应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取该图例键的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置图例键是否具有阴影效果
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置负值是否使用反色显示
    /// </summary>
    bool InvertIfNegative { get; set; }

    /// <summary>
    /// 获取或设置标记背景颜色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerBackgroundColor { get; set; }

    /// <summary>
    /// 获取或设置标记背景颜色索引
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    XlColorIndex MarkerBackgroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置标记前景颜色
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color MarkerForegroundColor { get; set; }

    /// <summary>
    /// 获取或设置标记前景颜色索引
    /// </summary>
    XlColorIndex MarkerForegroundColorIndex { get; set; }

    /// <summary>
    /// 获取或设置标记大小
    /// </summary>
    int MarkerSize { get; set; }

    /// <summary>
    /// 获取或设置标记样式
    /// </summary>
    XlMarkerStyle MarkerStyle { get; set; }

    /// <summary>
    /// 获取或设置图片类型
    /// </summary>
    [ConvertInt]
    [ComPropertyWrap(NeedConvert = true)]
    XlChartPictureType PictureType { get; set; }

    /// <summary>
    /// 获取或设置图片单位大小（新版属性）
    /// </summary>
    double PictureUnit2 { get; set; }

    /// <summary>
    /// 获取或设置图片单位大小（旧版属性）
    /// </summary>
    double PictureUnit { get; set; }

    /// <summary>
    /// 获取或设置线条是否平滑
    /// </summary>
    bool Smooth { get; set; }

    /// <summary>
    /// 获取图例键左侧位置（单位：磅）
    /// </summary>
    double Left { get; }

    /// <summary>
    /// 获取图例键顶部位置（单位：磅）
    /// </summary>
    double Top { get; }

    /// <summary>
    /// 获取图例键宽度（单位：磅）
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取图例键高度（单位：磅）
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取图例键内部区域格式
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取图例键填充格式
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取图例键边框格式
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取图例键的整体格式设置
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 清除图例键的所有格式设置
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? ClearFormats();

    /// <summary>
    /// 删除图例键
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? Delete();

    /// <summary>
    /// 选择图例键
    /// </summary>
    /// <returns>操作结果对象</returns>
    object? Select();
}