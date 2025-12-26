//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using Microsoft.Office.Interop.Excel;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Axis 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Axis 的安全访问和操作
/// 代表图表中的单个坐标轴（如类别轴、数值轴）
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelAxis : IOfficeObject<IExcelAxis>, IDisposable
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
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    #region 坐标轴属性
    /// <summary>
    /// 获取或设置坐标轴是否有标题
    /// 对应 Axis.HasTitle 属性
    /// </summary>
    bool HasTitle { get; set; }

    /// <summary>
    /// 获取坐标轴的格式设置对象
    /// 对应 Axis.Format 属性
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取或设置坐标轴的类型
    /// 对应 Axis.Type 属性 (使用 XlAxisType 枚举对应的 int 值)
    /// </summary>
    XlAxisType Type { get; set; }


    /// <summary>
    /// 获取或设置时间轴的基本单位
    /// 对应 Axis.BaseUnit 属性
    /// </summary>
    XlTimeUnit BaseUnit { get; set; }

    /// <summary>
    /// 获取或设置时间基准单位是否自动确定
    /// 对应 Axis.BaseUnitIsAuto 属性
    /// </summary>
    bool BaseUnitIsAuto { get; set; }

    /// <summary>
    /// 获取或设置主要单位比例
    /// 对应 Axis.MajorUnitScale 属性
    /// </summary>
    XlTimeUnit MajorUnitScale { get; set; }

    /// <summary>
    /// 获取或设置次要单位比例
    /// 对应 Axis.MinorUnitScale 属性
    /// </summary>
    XlTimeUnit MinorUnitScale { get; set; }

    /// <summary>
    /// 获取或设置分类轴的类型
    /// 对应 Axis.CategoryType 属性
    /// </summary>
    XlCategoryType CategoryType { get; set; }

    /// <summary>
    /// 获取或设置显示单位
    /// 对应 Axis.DisplayUnit 属性
    /// </summary>
    XlDisplayUnit DisplayUnit { get; set; }

    /// <summary>
    /// 获取或设置自定义显示单位值
    /// 对应 Axis.DisplayUnitCustom 属性
    /// </summary>
    double DisplayUnitCustom { get; set; }

    /// <summary>
    /// 获取或设置是否具有显示单位标签
    /// 对应 Axis.HasDisplayUnitLabel 属性
    /// </summary>
    bool HasDisplayUnitLabel { get; set; }

    /// <summary>
    /// 获取或设置是否具有主要网格线
    /// 对应 Axis.HasMajorGridlines 属性
    /// </summary>
    bool HasMajorGridlines { get; set; }

    /// <summary>
    /// 获取或设置是否具有次要网格线
    /// 对应 Axis.HasMinorGridlines 属性
    /// </summary>
    bool HasMinorGridlines { get; set; }

    /// <summary>
    /// 获取显示单位标签对象
    /// 对应 Axis.DisplayUnitLabel 属性
    /// </summary>
    IExcelDisplayUnitLabel? DisplayUnitLabel { get; }

    /// <summary>
    /// 获取坐标轴左边距位置
    /// 对应 Axis.Left 属性
    /// </summary>
    double Left { get; }

    /// <summary>
    /// 获取坐标轴上边距位置
    /// 对应 Axis.Top 属性
    /// </summary>
    double Top { get; }

    /// <summary>
    /// 获取坐标轴宽度
    /// 对应 Axis.Width 属性
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取坐标轴高度
    /// 对应 Axis.Height 属性
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取或设置刻度标签间距是否自动确定
    /// 对应 Axis.TickLabelSpacingIsAuto 属性
    /// </summary>
    bool TickLabelSpacingIsAuto { get; set; }

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
    XlTickMark MajorTickMark { get; set; }

    /// <summary>
    /// 获取或设置坐标轴次要刻度线的类型
    /// </summary>
    XlTickMark MinorTickMark { get; set; }

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

    /// <summary>
    /// 获取或设置分类轴的分类名称集合
    /// 对应 Axis.CategoryNames 属性
    /// </summary>
    object CategoryNames { get; set; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取坐标轴的边框格式设置对象
    /// 对应 Axis.Border 属性
    /// </summary>
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
    /// 删除坐标轴
    /// </summary>
    void Delete();

    object Select();

    #endregion
}
