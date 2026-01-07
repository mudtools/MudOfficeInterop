//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 使用 Microsoft.Office.Interop.Word.Shape 或 Microsoft.Office.Interop.Word.ShapeRange 对象的 HeightRelative 属性指定的值来指定形状的相对高度
/// </summary>
public enum WdRelativeVerticalSize
{
    /// <summary>
    /// 高度相对于左边距和右边距之间的空间
    /// </summary>
    wdRelativeVerticalSizeMargin,

    /// <summary>
    /// 高度相对于页面的高度
    /// </summary>
    wdRelativeVerticalSizePage,

    /// <summary>
    /// 高度相对于上边距的大小
    /// </summary>
    wdRelativeVerticalSizeTopMarginArea,

    /// <summary>
    /// 高度相对于下边距的大小
    /// </summary>
    wdRelativeVerticalSizeBottomMarginArea,

    /// <summary>
    /// 高度相对于内侧边距的大小——奇数页相对于上边距的大小，偶数页相对于下边距的大小
    /// </summary>
    wdRelativeVerticalSizeInnerMarginArea,

    /// <summary>
    /// 高度相对于外侧边距的大小——奇数页相对于下边距的大小，偶数页相对于上边距的大小
    /// </summary>
    wdRelativeVerticalSizeOuterMarginArea
}