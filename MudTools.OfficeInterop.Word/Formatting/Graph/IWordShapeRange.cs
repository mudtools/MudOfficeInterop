//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档中形状范围（ShapeRange）的封装接口。
/// </summary>
public interface IWordShapeRange : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取形状范围中的形状数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取形状。
    /// </summary>
    /// <param name="index">索引（从 1 开始）。</param>
    IWordShape this[object index] { get; }

    /// <summary>
    /// 获取或设置形状范围的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状范围的左边缘位置（相对于文档左边缘）。
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取形状范围的上边缘位置（相对于文档上边缘）。
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取形状范围的宽度。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取形状范围的高度。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置形状范围的水平对齐方式。
    /// </summary>
    WdShapePosition HorizontalFlip { get; }

    /// <summary>
    /// 获取或设置形状范围的垂直对齐方式。
    /// </summary>
    WdShapePosition VerticalFlip { get; }

    /// <summary>
    /// 获取形状范围的 Z 轴顺序。
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取形状范围的文本框架。
    /// </summary>
    IWordTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取形状范围的填充格式。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取形状范围的线条格式。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取形状范围的阴影格式。
    /// </summary>
    IWordShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取形状范围的三维格式。
    /// </summary>
    IWordThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取形状范围的调整选项。
    /// </summary>
    IWordAdjustments? Adjustments { get; }

    /// <summary>
    /// 获取形状范围的自动调整设置。
    /// </summary>
    MsoAutoShapeType AutoShapeType { get; set; }

    /// <summary>
    /// 获取形状范围的锁定锚点。
    /// </summary>
    IWordRange Anchor { get; }

    /// <summary>
    /// 获取形状范围的水平对齐相对于页面的设置。
    /// </summary>
    WdRelativeHorizontalPosition RelativeHorizontalPosition { get; set; }

    /// <summary>
    /// 获取形状范围的垂直对齐相对于页面的设置。
    /// </summary>
    WdRelativeVerticalPosition RelativeVerticalPosition { get; set; }

    /// <summary>
    /// 获取形状范围的布局方式。
    /// </summary>
    int LayoutInCell { get; set; }

    /// <summary>
    /// 删除形状范围中的所有形状。
    /// </summary>
    void Delete();

    void Align(MsoAlignCmd alignCmd, int relativeTo);

    void Apply();

    IWordShapeRange? Duplicate();

    /// <summary>
    /// 选择形状范围。
    /// </summary>
    void Select();

    /// <summary>
    /// 将形状范围移动到指定的 Z 轴顺序位置。
    /// </summary>
    /// <param name="zOrderCmd">Z 轴顺序命令。</param>
    void ZOrder(MsoZOrderCmd zOrderCmd);

    /// <summary>
    /// 组合形状范围中的形状。
    /// </summary>
    /// <returns>组合后的形状。</returns>
    IWordShape? Group();

    /// <summary>
    /// 取消组合形状范围中的形状。
    /// </summary>
    /// <returns>取消组合后的形状范围。</returns>
    IWordShapeRange? Ungroup();

    /// <summary>
    /// 重新排列形状范围中的形状。
    /// </summary>
    void Distribute(MsoDistributeCmd distributeCmd, int relativeTo);

    /// <summary>
    /// 转换为图片。
    /// </summary>
    void ConvertToInlineShape();

    /// <summary>
    /// 转换为浮动形状。
    /// </summary>
    void ConvertToFrame();
}