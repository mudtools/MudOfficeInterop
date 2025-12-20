//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 文档中图表的图例项接口
/// </summary>
/// <remarks>
/// 该接口封装了 Microsoft Word 图表中单个图例项的相关属性和方法，
/// 提供对图例项位置、格式和相关对象的访问能力。
/// </remarks>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLegendEntry : IDisposable
{
    /// <summary>
    /// 获取与该图例项关联的应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取该图例项的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取图例项在集合中的索引值
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取图例项左侧边界的坐标值（单位：磅）
    /// </summary>
    double Left { get; }

    /// <summary>
    /// 获取图例项顶部边界的坐标值（单位：磅）
    /// </summary>
    double Top { get; }

    /// <summary>
    /// 获取图例项的宽度（单位：磅）
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取图例项的高度（单位：磅）
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取或设置图例项中文本的自动缩放字体功能
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取图例项的字体格式设置
    /// </summary>
    IWordChartFont? Font { get; }

    /// <summary>
    /// 获取图例项的格式属性
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 获取与图例项相关联的图例键
    /// </summary>
    IWordLegendKey? LegendKey { get; }

    /// <summary>
    /// 删除图例项
    /// </summary>
    /// <returns>操作结果对象</returns>
    object Delete();

    /// <summary>
    /// 选择图例项
    /// </summary>
    /// <returns>操作结果对象</returns>
    object Select();
}