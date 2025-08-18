//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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
public interface IExcelChartFormat : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取 ChartFormat 对象的父对象
    /// 父对象通常是 ChartArea, PlotArea, Series 等图表元素
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取 ChartFormat 对象所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 格式设置
    /// <summary>
    /// 获取图表元素的填充格式对象
    /// 对应 ChartFormat.Fill 属性
    /// </summary>
    IExcelFillFormat Fill { get; }

    /// <summary>
    /// 获取图表元素的边框线条格式对象
    /// 对应 ChartFormat.Line 属性
    /// </summary>
    IExcelLine Line { get; } // 假设 IExcelLine 已定义

    // 注意：ChartFormat 还可能包含其他属性，如 Glow, Shadow, SoftEdge, TextFrame2 等，
    // 它们通常用于更高级的形状格式设置。可以根据需要继续扩展此接口。
    // IExcelGlow Glow { get; }
    // IExcelShadow Shadow { get; }
    // IExcelSoftEdge SoftEdge { get; }
    // MsExcel.TextFrame2 TextFrame2 { get; } // 直接暴露或进一步封装

    #endregion  
}