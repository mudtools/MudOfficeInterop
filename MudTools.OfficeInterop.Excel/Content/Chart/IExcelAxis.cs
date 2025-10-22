//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Axis 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Axis 的安全访问和操作
/// 代表图表中的单个坐标轴（如类别轴、数值轴）
/// </summary>
public interface IExcelAxis : IDisposable
{
    #region 基础属性  

    /// <summary>
    /// 获取坐标轴的父对象（通常是 Axes 集合）
    /// 对应 Axis.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取坐标轴所在的 Application 对象
    /// 对应 Axis.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 坐标轴属性
    bool HasTitle { get; set; }

    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取或设置坐标轴的类型
    /// 对应 Axis.Type 属性 (使用 XlAxisType 枚举对应的 int 值)
    /// </summary>
    XlAxisType Type { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的分组
    /// 对应 Axis.AxisGroup 属性 (使用 XlAxisGroup 枚举对应的 int 值)
    /// </summary>
    XlAxisGroup AxisGroup { get; }

    /// <summary>
    /// 获取或设置坐标轴标题
    /// 对应 Axis.HasTitle 和 Axis.AxisTitle 属性
    /// </summary>
    IExcelAxisTitle? AxisTitle { get; }

    /// <summary>
    /// 获取或设置坐标轴的位置类型（自动、最大值、最小值等）
    /// 对应 Axis.Crosses 属性 (使用 XlAxisCrosses 枚举对应的 int 值)
    /// </summary>
    XlAxisCrosses Crosses { get; set; }

    /// <summary>
    /// 获取或设置坐标轴在指定数值处穿过另一轴
    /// 对应 Axis.CrossesAt 属性
    /// </summary>
    double CrossesAt { get; set; }

    /// <summary>
    /// 获取或设置坐标轴是否在刻度线之间（类别轴）
    /// 对应 Axis.AxisBetweenCategories 属性
    /// </summary>
    bool AxisBetweenCategories { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的最小值
    /// 对应 Axis.MinimumScale 属性
    /// </summary>
    double MinimumScale { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的最大值
    /// 对应 Axis.MaximumScale 属性
    /// </summary>
    double MaximumScale { get; set; }

    /// <summary>
    /// 获取或设置坐标轴主要刻度单位
    /// 对应 Axis.MajorUnit 属性
    /// </summary>
    double MajorUnit { get; set; }

    /// <summary>
    /// 获取或设置坐标轴次要刻度单位
    /// 对应 Axis.MinorUnit 属性
    /// </summary>
    double MinorUnit { get; set; }

    /// <summary>
    /// 获取或设置坐标轴刻度线的类型（无、内部、外部、交叉）
    /// 对应 Axis.TickMarkSpacing 属性 (简化处理)
    /// 或更精确地使用 MajorTickMark / MinorTickMark (使用 XlTickMark 枚举对应的 int 值)
    /// </summary>
    XlTickMark MajorTickMark { get; set; } // 使用 XlTickMark
    /// <summary>
    /// 获取或设置坐标轴次要刻度线的类型
    /// </summary>
    XlTickMark MinorTickMark { get; set; } // 使用 XlTickMark

    /// <summary>
    /// 获取或设置坐标轴标签的位置（高、低、下一刻度线）
    /// 对应 Axis.TickLabelPosition 属性 (使用 XlTickLabelPosition 枚举对应的 int 值)
    /// </summary>
    XlTickLabelPosition TickLabelPosition { get; set; }

    /// <summary>
    /// 获取或设置坐标轴是否反转刻度值
    /// 对应 Axis.ReversePlotOrder 属性
    /// </summary>
    bool ReversePlotOrder { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的对数刻度底数（如果适用）
    /// 对应 Axis.LogBase 属性
    /// </summary>
    double LogBase { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的主要单位是否自动确定
    /// 对应 Axis.MajorUnitIsAuto 属性
    /// </summary>
    bool MajorUnitIsAuto { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的次要单位是否自动确定
    /// 对应 Axis.MinorUnitIsAuto 属性
    /// </summary>
    bool MinorUnitIsAuto { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的最小刻度值是否自动确定
    /// 对应 Axis.MinimumScaleIsAuto 属性
    /// </summary>
    bool MinimumScaleIsAuto { get; set; }

    /// <summary>
    /// 获取或设置坐标轴的最大刻度值是否自动确定
    /// 对应 Axis.MaximumScaleIsAuto 属性
    /// </summary>
    bool MaximumScaleIsAuto { get; set; }

    #endregion

    #region 格式设置
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取坐标轴刻度线标签对象
    /// 对应 Axis.TickLabels 属性
    /// </summary>
    IExcelTickLabels? TickLabels { get; }

    /// <summary>
    /// 获取坐标轴的主要网格线对象
    /// 对应 Axis.MajorGridlines 属性
    /// </summary>
    IExcelGridlines? MajorGridlines { get; }

    /// <summary>
    /// 获取坐标轴的次要网格线对象
    /// 对应 Axis.MinorGridlines 属性
    /// </summary>
    IExcelGridlines? MinorGridlines { get; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除坐标轴（通常不直接删除，而是通过图表设置隐藏）
    /// </summary>
    void Delete();

    #endregion
}
