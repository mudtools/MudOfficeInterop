//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

using System;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shape 的接口，用于操作文档中的形状对象。
/// </summary>
public interface IWordShape : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取形状的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状的类型。
    /// </summary>
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取或设置形状的左边距（相对于页面）。
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置形状的上边距（相对于页面）。
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置形状的高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置形状的相对水平位置。
    /// </summary>
    WdRelativeHorizontalPosition RelativeHorizontalPosition { get; set; }

    /// <summary>
    /// 获取或设置形状的相对垂直位置。
    /// </summary>
    WdRelativeVerticalPosition RelativeVerticalPosition { get; set; }

    /// <summary>
    /// 获取形状的文本框对象。
    /// </summary>
    IWordTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取形状的填充格式。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取形状的线条格式。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取形状的阴影格式。
    /// </summary>
    IWordShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取形状的三维格式。
    /// </summary>
    IWordThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取形状的链接格式。
    /// </summary>
    IWordLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取形状的OLE格式。
    /// </summary>
    IWordOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取或设置图片的柔化边缘格式。
    /// </summary>
    IWordSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取或设置图片的光泽格式。
    /// </summary>
    IWordReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取或设置图片的反射格式。
    /// </summary>
    IWordGlowFormat? Glow { get; }

    /// <summary>
    /// 获取形状的文本环绕格式。
    /// </summary>
    IWordWrapFormat? WrapFormat { get; }

    /// <summary>
    /// 获取或设置形状是否锁定纵横比。
    /// </summary>
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取或设置形状是否可旋转。
    /// </summary>
    float Rotation { get; set; }

    /// <summary>
    /// 获取或设置形状的替代文本。
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取或设置形状的Z轴顺序。
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取形状是否为浮动形状。
    /// </summary>
    bool IsFloating { get; }

    /// <summary>
    /// 获取形状是否为内联形状。
    /// </summary>
    bool IsInline { get; }

    /// <summary>
    /// 获取形状所在的范围。
    /// </summary>
    IWordRange? Anchor { get; }

    /// <summary>
    /// 删除形状。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择形状。
    /// </summary>
    void Select();

    /// <summary>
    /// 将形状移动到指定的Z轴位置。
    /// </summary>
    /// <param name="position">Z轴位置。</param>
    void ZOrder(MsoZOrderCmd position);

    /// <summary>
    /// 调整形状大小。
    /// </summary>
    void ScaleHeight(float Factor, bool RelativeToOriginalSize, MsoScaleFrom Scale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 调整形状宽度。
    /// </summary>
    void ScaleWidth(float Factor, bool RelativeToOriginalSize, MsoScaleFrom Scale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 旋转形状。
    /// </summary>
    /// <param name="increment">旋转角度增量。</param>
    void IncrementRotation(float increment);

    /// <summary>
    /// 水平翻转形状。
    /// </summary>
    void FlipHorizontal();

    /// <summary>
    /// 垂直翻转形状。
    /// </summary>
    void FlipVertical();

    /// <summary>
    /// 将形状转换为内联形状。
    /// </summary>
    /// <returns>转换后的内联形状。</returns>
    IWordInlineShape? ConvertToInlineShape();

    /// <summary>
    /// 将内联形状转换为浮动形状。
    /// </summary>
    /// <returns>转换后的形状。</returns>
    IWordFrame? ConvertToFrame();

    /// <summary>
    /// 获取形状的图表对象（如果形状是图表）。
    /// </summary>
    IWordChart? Chart { get; }

    /// <summary>
    /// 获取形状的SmartArt对象（如果形状是SmartArt）。
    /// </summary>
    IOfficeSmartArt? SmartArt { get; }


}