//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ChartFormat 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ChartFormat 的安全访问和操作
/// ChartFormat 对象包含图表元素（如 ChartArea, PlotArea, Series 等）的通用格式属性
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelChartFormat : IOfficeObject<IExcelChartFormat, MsExcel.ChartFormat>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 ChartFormat 对象的父对象
    /// 父对象通常是 ChartArea, PlotArea, Series 等图表元素
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取 ChartFormat 对象所在的 Application 对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }
    #endregion

    #region 格式设置

    /// <summary>
    /// 获取图表元素的填充格式对象
    /// </summary>
    IExcelFillFormat? Fill { get; }

    /// <summary>
    /// 获取线条格式对象
    /// </summary>
    IExcelLineFormat? Line { get; }

    /// <summary>
    /// 获取阴影格式对象
    /// </summary>
    IExcelShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取三维格式对象
    /// </summary>
    IExcelThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取调整点集合对象（用于自定义形状）
    /// </summary>
    IExcelAdjustments? Adjustments { get; }

    /// <summary>
    /// 获取图片格式对象
    /// </summary>
    IExcelPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取柔化边缘效果对象（Office 2010+）
    /// </summary>
    IOfficeSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取发光效果对象（Office 2010+）
    /// </summary>
    IOfficeGlowFormat? Glow { get; }

    /// <summary>
    /// 获取文本框架对象（Office 2010+）
    /// 用于设置和操作图表元素中文本框的格式属性
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    IExcelTextFrame2 TextFrame2 { get; }

    /// <summary>
    /// 获取或设置自动形状类型
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoShapeType AutoShapeType { get; set; }
    #endregion
}