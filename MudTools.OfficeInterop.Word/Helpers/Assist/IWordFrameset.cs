//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示框架页的整个页面或框架页上的单个框架。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFrameset : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 返回表示框架页上指定 Frameset 对象父级的 Frameset 对象。
    /// </summary>
    IWordFrameset? ParentFrameset { get; }

    /// <summary>
    /// 返回 Microsoft.Office.Interop.Word.Frameset 对象的类型。
    /// </summary>
    WdFramesetType Type { get; }

    /// <summary>
    /// 返回或设置指定 Microsoft.Office.Interop.Word.Frameset 对象的宽度类型。
    /// </summary>
    WdFramesetSizeType WidthType { get; set; }

    /// <summary>
    /// 返回或设置指定框架页上框架的高度类型。
    /// </summary>
    WdFramesetSizeType HeightType { get; set; }

    /// <summary>
    /// 返回或设置指定 Microsoft.Office.Interop.Word.Frameset 对象的宽度。
    /// </summary>
    int Width { get; set; }

    /// <summary>
    /// 返回或设置指定 Microsoft.Office.Interop.Word.Frameset 对象的高度。
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// 返回与指定 Frameset 对象关联的子 Frameset 对象数量。
    /// </summary>
    int ChildFramesetCount { get; }

    /// <summary>
    /// 返回表示指定子 Frameset 对象的 Frameset 对象。
    /// </summary>
    /// <param name="index">必需 Integer。指定框架的索引号。</param>
    /// <returns>Microsoft.Office.Interop.Word.Frameset</returns>
    [MethodIndex]
    IWordFrameset? ChildFramesetItem(int index);

    /// <summary>
    /// 返回或设置指定框架页上框架周围边框的宽度（以磅为单位）。
    /// </summary>
    float FramesetBorderWidth { get; set; }

    /// <summary>
    /// 返回或设置指定框架页上框架边框的颜色。可以是任何 Microsoft.Office.Interop.Word.WdColor 常量。
    /// </summary>
    WdColor FramesetBorderColor { get; set; }

    /// <summary>
    /// 返回或设置在 Web 浏览器中查看框架页时，指定框架何时可提供滚动条。
    /// </summary>
    WdScrollbarType FrameScrollbarType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示当在 Web 浏览器中查看框架页时，用户是否可以调整指定框架的大小。
    /// </summary>
    bool FrameResizable { get; set; }

    /// <summary>
    /// 返回或设置框架页上指定框架的名称。
    /// </summary>
    string FrameName { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否显示指定框架页上的框架边框。
    /// </summary>
    bool FrameDisplayBorders { get; set; }

    /// <summary>
    /// 返回或设置在打开框架页时要在指定框架中显示的网页或其他文档。
    /// </summary>
    string FrameDefaultURL { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示由 Microsoft.Office.Interop.Word.Frameset.FrameDefaultURL 属性指定的网页或其他文档是否是外部文件，Microsoft Word 仅从指定框架保持链接。
    /// </summary>
    bool FrameLinkToFile { get; set; }

    /// <summary>
    /// 向框架页添加新框架。
    /// </summary>
    /// <param name="where">必需 Microsoft.Office.Interop.Word.WdFramesetNewFrameLocation。设置相对于指定框架的新框架要添加的位置。</param>
    /// <returns>Microsoft.Office.Interop.Word.Frameset</returns>
    IWordFrameset? AddNewFrame(WdFramesetNewFrameLocation where);

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

}