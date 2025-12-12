//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示Office线条格式对象的接口
/// 封装了Microsoft.Office.Core.LineFormat COM对象
/// 用于处理线条和形状边框的格式设置
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeLineFormat : IDisposable
{
    /// <summary>
    /// 获取或设置线条的可见性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置线条的粗细（以磅为单位）
    /// </summary>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置线条的虚线样式
    /// </summary>
    MsoLineDashStyle DashStyle { get; set; }

    /// <summary>
    /// 获取或设置线条样式
    /// </summary>
    MsoLineStyle Style { get; set; }

    /// <summary>
    /// 获取线条的前景颜色格式
    /// </summary>
    IOfficeColorFormat ForeColor { get; }

    /// <summary>
    /// 获取线条的背景颜色格式（用于图案线条）
    /// </summary>
    IOfficeColorFormat BackColor { get; }

    /// <summary>
    /// 获取或设置线条的透明度（0-1之间，0为完全不透明，1为完全透明）
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置线条起点箭头head的长度
    /// </summary>
    MsoArrowheadLength BeginArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条起点箭头head的样式
    /// </summary>
    MsoArrowheadStyle BeginArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条起点箭头head的宽度
    /// </summary>
    MsoArrowheadWidth BeginArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条终点箭头head的长度
    /// </summary>
    MsoArrowheadLength EndArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条终点箭头head的样式
    /// </summary>
    MsoArrowheadStyle EndArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条终点箭头head的宽度
    /// </summary>
    MsoArrowheadWidth EndArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的图案类型（用于图案线条）
    /// </summary>
    MsoPatternType Pattern { get; set; }
}