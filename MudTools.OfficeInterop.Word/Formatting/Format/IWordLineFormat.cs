//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Core.LineFormat 的接口，用于操作线条格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordLineFormat : IOfficeObject<IWordLineFormat, MsWord.LineFormat>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置线条的前景颜色格式。
    /// </summary>
    IWordColorFormat? ForeColor { get; }

    /// <summary>
    /// 获取或设置线条的背景颜色格式。
    /// </summary>
    IWordColorFormat? BackColor { get; }

    /// <summary>
    /// 获取或设置线条的透明度（0.0到1.0之间）。
    /// </summary>
    float Transparency { get; set; }

    /// <summary>
    /// 获取或设置线条是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置线条的粗细（磅）。
    /// </summary>
    float Weight { get; set; }

    /// <summary>
    /// 获取或设置线条样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLineStyle Style { get; set; }

    /// <summary>
    /// 获取或设置线条的虚线样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoLineDashStyle DashStyle { get; set; }

    /// <summary>
    /// 获取或设置线条的端点样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadStyle BeginArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条的起始箭头宽度。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadWidth BeginArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的起始箭头长度。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadLength BeginArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置线条的结束端点样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadStyle EndArrowheadStyle { get; set; }

    /// <summary>
    /// 获取或设置线条的结束箭头宽度。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadWidth EndArrowheadWidth { get; set; }

    /// <summary>
    /// 获取或设置线条的结束箭头长度。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoArrowheadLength EndArrowheadLength { get; set; }

    /// <summary>
    /// 获取或设置图案类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPatternType Pattern { get; set; }

    /// <summary>
    /// 获取或设置线条是否使用插入笔。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool InsetPen { get; set; }
}