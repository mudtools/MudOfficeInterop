//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Line (边框线条) 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.LineFormat 的安全访问和操作
/// 用于设置形状或图表元素的边框线条
/// </summary>
public interface IExcelLine : IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取线条所在的父对象
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取线条对象所在的 Application 对象
    /// </summary>
    IExcelApplication Application { get; } // 假设 IExcelApplication 已定义
    #endregion

    #region 线条属性
    /// <summary>
    /// 获取或设置线条的颜色 (RGB 颜色值)
    /// 对应 LineFormat.ForeColor
    /// </summary>
    int Color { get; set; }

    /// <summary>
    /// 获取或设置线条的样式
    /// 对应 LineFormat.Style (使用 MsoLineStyle 枚举对应的 int 值)
    /// </summary>
    MsoLineStyle Style { get; set; }

    /// <summary>
    /// 获取或设置线条的粗细 (磅)
    /// 对应 LineFormat.Weight
    /// </summary>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置线条是否可见
    /// 对应 LineFormat.Visible (使用 MsoTriState 枚举对应的 int 值)
    /// </summary>
    bool Visible { get; set; } // Using int for MsoTriState

    /// <summary>
    /// 获取或设置线条的透明度 (0.0 = 不透明, 1.0 = 完全透明)
    /// 对应 LineFormat.Transparency
    /// </summary>
    float Transparency { get; set; }

    // --- 高级线条属性 (占位符) ---
    // /// <summary>
    // /// 获取或设置虚线样式 (如果 Style 支持)
    // /// </summary>
    // int DashStyle { get; set; } // MsoLineDashStyle
    // /// <summary>
    // /// 获取或设置线条端点类型
    // /// </summary>
    // int EndCap { get; set; } // MsoLineEndCap
    // /// <summary>
    // /// 获取或设置起始箭头样式
    // /// </summary>
    // int BeginArrowheadStyle { get; set; } // MsoArrowheadStyle
    // /// <summary>
    // /// 获取或设置结束箭头样式
    // /// </summary>
    // int EndArrowheadStyle { get; set; } // MsoArrowheadStyle
    #endregion 
}