//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表轴的封装接口。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordAxis : IOfficeObject<IWordAxis>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取轴类型。
    /// </summary>
    XlAxisType Type { get; }

    /// <summary>
    /// 获取轴组。
    /// </summary>
    XlAxisGroup AxisGroup { get; }

    /// <summary>
    /// 获取或设置轴标题。
    /// </summary>
    IWordAxisTitle? AxisTitle { get; }

    /// <summary>
    /// 获取或设置是否显示轴标题。
    /// </summary>
    bool HasTitle { get; set; }

    /// <summary>
    /// 获取或设置是否显示刻度线标签。
    /// </summary>
    bool HasMajorGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否显示次刻度线标签。
    /// </summary>
    bool HasMinorGridlines { get; set; }

    /// <summary>
    /// 获取或设置刻度线标签位置。
    /// </summary>
    XlTickLabelPosition TickLabelPosition { get; set; }

    /// <summary>
    /// 获取或设置主刻度线类型。
    /// </summary>
    XlTickMark MajorTickMark { get; set; }

    /// <summary>
    /// 获取或设置次刻度线类型。
    /// </summary>
    XlTickMark MinorTickMark { get; set; }

    /// <summary>
    /// 获取或设置轴刻度线间距。
    /// </summary>
    double MajorUnit { get; set; }

    /// <summary>
    /// 获取或设置是否自动设置主刻度线间距。
    /// </summary>
    bool MajorUnitIsAuto { get; set; }

    /// <summary>
    /// 获取或设置次刻度线间距。
    /// </summary>
    double MinorUnit { get; set; }

    /// <summary>
    /// 获取或设置是否自动设置次刻度线间距。
    /// </summary>
    bool MinorUnitIsAuto { get; set; }

    /// <summary>
    /// 获取或设置轴最小值。
    /// </summary>
    double MinimumScale { get; set; }

    /// <summary>
    /// 获取或设置是否自动设置轴最小值。
    /// </summary>
    bool MinimumScaleIsAuto { get; set; }

    /// <summary>
    /// 获取或设置轴最大值。
    /// </summary>
    double MaximumScale { get; set; }

    /// <summary>
    /// 获取或设置是否自动设置轴最大值。
    /// </summary>
    bool MaximumScaleIsAuto { get; set; }

    /// <summary>
    /// 获取或设置轴交叉点。
    /// </summary>
    double CrossesAt { get; set; }

    /// <summary>
    /// 获取或设置轴交叉方式。
    /// </summary>
    XlAxisCrosses Crosses { get; set; }

    /// <summary>
    /// 获取或设置是否反转轴。
    /// </summary>
    bool ReversePlotOrder { get; set; }

    /// <summary>
    /// 获取或设置是否对数刻度。
    /// </summary>
    XlScaleType ScaleType { get; set; }

    /// <summary>
    /// 获取刻度线标签对象。
    /// </summary>
    IWordTickLabels? TickLabels { get; }

    /// <summary>
    /// 获取主网格线对象。
    /// </summary>
    IWordGridlines? MajorGridlines { get; }

    /// <summary>
    /// 获取次网格线对象。
    /// </summary>
    IWordGridlines? MinorGridlines { get; }

    /// <summary>
    /// 获取边框格式。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 选择轴。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除轴。
    /// </summary>
    void Delete();
}