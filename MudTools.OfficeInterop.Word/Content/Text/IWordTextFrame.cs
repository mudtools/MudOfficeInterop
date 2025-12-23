//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.TextFrame 的接口，用于操作文本框格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordTextFrame : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取文本框的文本范围。
    /// </summary>
    IWordRange? TextRange { get; }

    /// <summary>
    /// 获取或设置文本框的左边距（磅）。
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置文本框的右边距（磅）。
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置文本框的上边距（磅）。
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置文本框的下边距（磅）。
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取或设置文本框的水平对齐方式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoHorizontalAnchor HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框的垂直对齐方式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoVerticalAnchor VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框的自动调整大小方式。
    /// </summary>
    int AutoSize { get; set; }


    /// <summary>
    /// 获取或设置文本框的路径格式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPathFormat PathFormat { get; set; }

    /// <summary>
    /// 获取或设置文本框的文本方向。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置文本框的变形格式，用于控制文本的变形效果（如弯曲、倾斜等）。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoWarpFormat WarpFormat { get; set; }

    /// <summary>
    /// 获取或设置是否禁用文本旋转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool NoTextRotation { get; set; }

    /// <summary>
    /// 获取或设置文本框的自动换行设置（1表示启用，0表示禁用）。
    /// </summary>
    int WordWrap { get; set; }

    /// <summary>
    /// 获取文本框中是否包含文本（1表示有文本，0表示无文本）。
    /// </summary>
    int HasText { get; }

    /// <summary>
    /// 获取文本框是否内容溢出。
    /// </summary>
    bool Overflowing { get; }


    /// <summary>
    /// 获取文本框的列对象。
    /// </summary>
    IOfficeTextColumn2? Column { get; }

    /// <summary>
    /// 获取文本框的三维格式对象。
    /// </summary>
    IWordThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取前一个文本框（链接的文本框链中的前一个）。
    /// </summary>
    IWordTextFrame? Previous { get; }

    /// <summary>
    /// 获取下一个文本框（链接的文本框链中的下一个）。
    /// </summary>
    IWordTextFrame? Next { get; }

    /// <summary>
    /// 获取包含此文本框的范围对象。
    /// </summary>
    IWordRange? ContainingRange { get; }

    /// <summary>
    /// 获取此文本框的父级形状对象。
    /// </summary>
    IWordShape? Parent { get; }

    /// <summary>
    /// 删除文本框中的所有文本内容。
    /// </summary>
    void DeleteText();

    /// <summary>
    /// 验证目标文本框是否可以作为链接目标。
    /// </summary>
    /// <param name="targetTextFrame">目标文本框对象</param>
    /// <returns>如果可以链接则返回true，否则返回false，null表示无法确定</returns>
    bool? ValidLinkTarget(IWordTextFrame targetTextFrame);

    /// <summary>
    /// 断开到下一个文本框的链接关系。
    /// </summary>
    void BreakForwardLink();

}