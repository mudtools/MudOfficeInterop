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
public interface IWordInlineShape : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取内联形状的类型。
    /// </summary>
    WdInlineShapeType Type { get; }

    /// <summary>
    /// 获取内联形状所在的范围。
    /// </summary>
    IWordRange Range { get; }

    IWordTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取内联形状的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取或设置内联形状的宽度（磅）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置内联形状的高度（磅）。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置内联形状是否锁定纵横比。
    /// </summary>
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取内联形状的OLE格式。
    /// </summary>
    IWordOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取内联形状的链接格式。
    /// </summary>
    IWordLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取内联形状的字段对象。
    /// </summary>
    IWordField? Field { get; }

    /// <summary>
    /// 获取内联形状的线条格式。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取内联形状的填充格式。
    /// </summary>
    IWordFillFormat? Fill { get; }

    IWordShadowFormat? Shadow { get; }


    /// <summary>
    /// 获取内联形状的图表对象。
    /// </summary>
    IWordChart? Chart { get; }

    /// <summary>
    /// 获取内联形状的SmartArt对象。
    /// </summary>
    //IWordSmartArt SmartArt { get; }

    /// <summary>
    /// 获取内联形状的图片格式。
    /// </summary>
    IWordPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取内联形状的图形对象。
    /// </summary>
    IWordGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取内联形状是否为图片类型。
    /// </summary>
    bool IsPicture { get; }

    /// <summary>
    /// 获取内联形状是否为OLE对象。
    /// </summary>
    bool IsOLEObject { get; }

    /// <summary>
    /// 获取内联形状是否为图表。
    /// </summary>
    bool IsChart { get; }

    /// <summary>
    /// 获取内联形状是否为第一个形状。
    /// </summary>
    bool IsFirst { get; }

    /// <summary>
    /// 获取内联形状是否为最后一个形状。
    /// </summary>
    bool IsLast { get; }

    /// <summary>
    /// 删除内联形状。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择内联形状。
    /// </summary>
    void Select();

    /// <summary>
    /// 复制内联形状。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切内联形状。
    /// </summary>
    void Cut();

    /// <summary>
    /// 调整内联形状大小。
    /// </summary>
    /// <param name="width">新宽度。</param>
    /// <param name="height">新高度。</param>
    /// <param name="scale">是否按比例缩放。</param>
    void ScaleSize(float width, float height, bool scale = true);

    /// <summary>
    /// 将内联形状转换为浮动形状。
    /// </summary>
    /// <returns>转换后的浮动形状。</returns>
    IWordShape? ConvertToShape();

    /// <summary>
    /// 设置内联形状大小。
    /// </summary>
    /// <param name="width">宽度。</param>
    /// <param name="height">高度。</param>
    void SetSize(float width, float height);

    /// <summary>
    /// 重置内联形状大小为原始大小。
    /// </summary>
    void ResetSize();

    /// <summary>
    /// 复制内联形状格式到另一个内联形状。
    /// </summary>
    /// <param name="targetInlineShape">目标内联形状。</param>
    void CopyTo(IWordInlineShape targetInlineShape);

    /// <summary>
    /// 重置内联形状格式为默认值。
    /// </summary>
    void Reset();


    /// <summary>
    /// 更新链接的内联形状。
    /// </summary>
    /// <returns>是否更新成功。</returns>
    bool Update();

    /// <summary>
    /// 断开链接的内联形状。
    /// </summary>
    /// <returns>是否断开成功。</returns>
    bool BreakLink();

    /// <summary>
    /// 获取内联形状的替代文本。
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取内联形状的标题。
    /// </summary>
    string Title { get; set; }
}