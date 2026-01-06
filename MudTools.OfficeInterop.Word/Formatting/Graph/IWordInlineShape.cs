//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.InlineShape 的接口，用于操作内联形状对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordInlineShape : IOfficeObject<IWordInlineShape, MsWord.InlineShape>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取内联形状的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置表示指定对象的所有边框的 Borders 集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取表示指定对象中包含的文档部分的 Range 对象。
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取表示链接到文件的指定内嵌形状的链接选项的 LinkFormat 对象。
    /// </summary>
    IWordLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取表示与指定形状关联的字段的 Field 对象。
    /// </summary>
    IWordField? Field { get; }

    /// <summary>
    /// 获取表示指定内嵌形状的 OLE 特性（链接除外）的 OLEFormat 对象。
    /// </summary>
    IWordOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取内嵌形状的类型。
    /// </summary>
    WdInlineShapeType Type { get; }

    /// <summary>
    /// 获取表示与指定内嵌形状对象关联的超链接的 Hyperlink 对象。
    /// </summary>
    IWordHyperlink? Hyperlink { get; }

    /// <summary>
    /// 获取或设置指定内嵌形状的高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置指定对象的宽度（以磅为单位）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 相对于原始大小缩放指定内嵌形状的高度。
    /// </summary>
    float ScaleHeight { get; set; }

    /// <summary>
    /// 相对于原始大小缩放指定内嵌形状的宽度。
    /// </summary>
    float ScaleWidth { get; set; }

    /// <summary>
    /// 确定调整大小时指定形状是否保持其原始比例。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取包含指定形状的线条格式属性的 LineFormat 对象。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取包含指定形状的填充格式属性的 FillFormat 对象。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取或设置包含指定对象的图片格式属性的 PictureFormat 对象。
    /// </summary>
    IWordPictureFormat? PictureFormat { get; set; }

    /// <summary>
    /// 获取包含指定内嵌形状对象的水平线格式的 HorizontalLineFormat 对象。
    /// </summary>
    IWordHorizontalLineFormat? HorizontalLineFormat { get; }

    /// <summary>
    /// 获取表示指定网页上的脚本或代码块的 Script 对象。
    /// </summary>
    IOfficeScript? Script { get; }

    /// <summary>
    /// 获取包含指定形状的文本效果格式属性的 TextEffectFormat 对象。
    /// </summary>
    IWordTextEffectFormat? TextEffect { get; set; }

    /// <summary>
    /// 获取或设置与网页中形状关联的替代文本。
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 确定内嵌形状对象是否为图片项目符号。
    /// </summary>
    bool IsPictureBullet { get; }

    /// <summary>
    /// 获取内嵌形状中分组在一起的形状集合。只读。
    /// </summary>
    IWordGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取一个值，指示指定的形状是否为图表。只读。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasChart { get; }

    /// <summary>
    /// 获取文档中内嵌形状集合内的图表。只读。
    /// </summary>
    IWordChart? Chart { get; }

    /// <summary>
    /// 获取形状的柔化边缘格式。只读。
    /// </summary>
    IWordSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取发光效果的格式属性。只读。
    /// </summary>
    IWordGlowFormat? Glow { get; }

    /// <summary>
    /// 获取形状的反射格式。只读。
    /// </summary>
    IWordReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取指定形状的阴影格式。只读。
    /// </summary>
    IWordShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取一个值，指示形状上是否存在 SmartArt 图表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasSmartArt { get; }

    /// <summary>
    /// 获取提供与指定内嵌形状关联的 SmartArt 一起工作的方式的 SmartArt 对象。
    /// </summary>
    IOfficeSmartArt? SmartArt { get; }

    /// <summary>
    /// 获取或设置包含指定内嵌形状的标题的值。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 激活指定的对象。
    /// </summary>
    void Activate();

    /// <summary>
    /// 删除对内嵌形状所做的更改。
    /// </summary>
    void Reset();

    /// <summary>
    /// 删除指定的对象。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择指定的对象。
    /// </summary>
    void Select();

    /// <summary>
    /// 将内嵌形状转换为自由浮动的形状，并返回表示新形状的 Shape 对象。
    /// </summary>
    /// <returns>表示新形状的 Shape 对象。</returns>
    IWordShape? ConvertToShape();
}