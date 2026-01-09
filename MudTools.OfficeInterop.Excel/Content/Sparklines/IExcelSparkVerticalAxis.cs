//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示一组迷你图的垂直轴设置。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSparkVerticalAxis : IOfficeObject<IExcelSparkVerticalAxis, MsExcel.SparkVerticalAxis>, IDisposable
{
    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置迷你图垂直轴的最小值相对于组中其他迷你图的缩放方式。
    /// </summary>
    XlSparkScale MinScaleType { get; set; }

    /// <summary>
    /// 获取或设置迷你图垂直轴的自定义最小值。
    /// </summary>
    object CustomMinScaleValue { get; set; }

    /// <summary>
    /// 获取或设置迷你图垂直轴的最大值相对于组中其他迷你图的缩放方式。
    /// </summary>
    XlSparkScale MaxScaleType { get; set; }

    /// <summary>
    /// 获取或设置迷你图垂直轴的自定义最大值。
    /// </summary>
    object CustomMaxScaleValue { get; set; }
}