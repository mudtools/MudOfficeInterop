//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Frame 的接口，用于操作框架格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFrame : IOfficeObject<IWordFrame, MsWord.Frame>, IDisposable
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
    /// 获取或设置确定框架高度的规则。
    /// </summary>
    WdFrameSizeRule HeightRule { get; set; }

    /// <summary>
    /// 获取或设置确定框架宽度的规则。
    /// </summary>
    WdFrameSizeRule WidthRule { get; set; }

    /// <summary>
    /// 获取或设置框架与周围文本之间的水平距离（以磅为单位）。
    /// </summary>
    float HorizontalDistanceFromText { get; set; }

    /// <summary>
    /// 获取或设置框架的高度（以磅为单位）。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置框架边缘与 RelativeHorizontalPosition 属性指定项之间的水平距离。
    /// </summary>
    float HorizontalPosition { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示框架是否锁定其定位点。
    /// 如果为 true，则框架的定位点不会随文本移动。
    /// </summary>
    bool LockAnchor { get; set; }

    /// <summary>
    /// 获取或设置框架水平位置的相对参照物。
    /// </summary>
    WdRelativeHorizontalPosition RelativeHorizontalPosition { get; set; }

    /// <summary>
    /// 获取或设置框架垂直位置的相对参照物。
    /// </summary>
    WdRelativeVerticalPosition RelativeVerticalPosition { get; set; }

    /// <summary>
    /// 获取或设置框架与周围文本之间的垂直距离（以磅为单位）。
    /// </summary>
    float VerticalDistanceFromText { get; set; }

    /// <summary>
    /// 获取或设置框架边缘与 RelativeVerticalPosition 属性指定项之间的垂直距离。
    /// </summary>
    float VerticalPosition { get; set; }

    /// <summary>
    /// 获取或设置框架的宽度（以磅为单位）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文档文本是否环绕框架。
    /// 如果为 true，则文本会环绕框架；否则，框架将独占一行。
    /// </summary>
    bool TextWrap { get; set; }

    /// <summary>
    /// 获取框架的底纹格式设置。
    /// </summary>
    IWordShading? Shading { get; }

    /// <summary>
    /// 获取或设置框架的所有边框。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取包含在框架内的文档部分的范围。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 删除框架。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选中框架。
    /// </summary>
    void Select();

    /// <summary>
    /// 将框架复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 将框架从文档中剪切到剪贴板。
    /// </summary>
    void Cut();
}