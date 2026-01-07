//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shape 的接口，用于操作文档中的形状对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordShape : IOfficeObject<IWordShape, MsWord.Shape>, IDisposable
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
    /// 获取包含自选图形或艺术字所有调整值的 Adjustments 对象。
    /// </summary>
    IWordAdjustments? Adjustments { get; }

    /// <summary>
    /// 获取或设置指定自选图形的形状类型（不能是线条或自由曲线）。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoShapeType AutoShapeType { get; set; }

    /// <summary>
    /// 获取包含图形标注格式属性的 CalloutFormat 对象。
    /// </summary>
    IWordCalloutFormat? Callout { get; }

    /// <summary>
    /// 获取包含图形填充格式属性的 FillFormat 对象。
    /// </summary>
    IWordFillFormat? Fill { get; }

    /// <summary>
    /// 获取表示组中各个图形的 GroupShapes 对象。
    /// </summary>
    IWordGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取或设置图形的高度（以磅为单位）。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取一个值，指示图形是否已水平翻转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HorizontalFlip { get; }

    /// <summary>
    /// 获取或设置图形的水平位置（以磅为单位）。
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取包含图形线条格式属性的 LineFormat 对象。
    /// </summary>
    IWordLineFormat? Line { get; }

    /// <summary>
    /// 获取或设置一个值，指示调整图形大小时是否保持原始比例。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取或设置图形的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取表示图形几何描述的 ShapeNodes 集合。
    /// </summary>
    IWordShapeNodes? Nodes { get; }

    /// <summary>
    /// 获取或设置图形围绕 Z 轴旋转的度数。
    /// </summary>
    float Rotation { get; set; }

    /// <summary>
    /// 获取包含图形图片格式属性的 PictureFormat 对象。
    /// </summary>
    IWordPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取表示图形阴影格式的 ShadowFormat 对象。
    /// </summary>
    IWordShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取包含图形文本效果格式属性的 TextEffectFormat 对象。
    /// </summary>
    IWordTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取包含图形文本的 TextFrame 对象。
    /// </summary>
    IWordTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取包含图形三维效果格式属性的 ThreeDFormat 对象。
    /// </summary>
    IWordThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取或设置图形的垂直位置（以磅为单位）。
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取图形的类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取一个值，指示图形是否已垂直翻转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool VerticalFlip { get; }

    /// <summary>
    /// 获取自由曲线的顶点坐标（以及贝塞尔曲线的控制点坐标）。
    /// </summary>
    object Vertices { get; }

    /// <summary>
    /// 获取或设置一个值，指示图形或应用于它的格式是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置图形的宽度（以磅为单位）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取图形在 Z 轴顺序中的位置。
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取与图形关联的超链接。
    /// </summary>
    IWordHyperlink? Hyperlink { get; }

    /// <summary>
    /// 获取或设置图形水平位置的相对参照物。
    /// </summary>
    WdRelativeHorizontalPosition RelativeHorizontalPosition { get; set; }

    /// <summary>
    /// 获取或设置图形垂直位置的相对参照物。
    /// </summary>
    WdRelativeVerticalPosition RelativeVerticalPosition { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示图形的定位点是否锁定到定位范围。
    /// </summary>
    int LockAnchor { get; set; }

    /// <summary>
    /// 获取包含图形文本环绕属性的 WrapFormat 对象。
    /// </summary>
    IWordWrapFormat? WrapFormat { get; }

    /// <summary>
    /// 获取表示图形 OLE 特性的 OLEFormat 对象（除链接外）。
    /// </summary>
    IWordOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取表示图形定位范围的 Range 对象。
    /// 所有图形都锚定到文本范围，但可以放置在包含锚点的页面的任何位置。
    /// </summary>
    IWordRange? Anchor { get; }

    /// <summary>
    /// 获取表示链接到文件的图形的链接选项的 LinkFormat 对象。
    /// </summary>
    IWordLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取或设置与网页中形状关联的替代文本。
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取表示网页上脚本或代码块的 Script 对象。
    /// </summary>
    IOfficeScript? Script { get; }

    /// <summary>
    /// 获取一个值，指示形状是否为子形状。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Child { get; }

    /// <summary>
    /// 获取表示子形状的公共父形状的 Shape 对象。
    /// </summary>
    IWordShape? ParentGroup { get; }

    /// <summary>
    /// 获取表示绘图画布中形状集合的 CanvasShapes 对象。
    /// </summary>
    IWordCanvasShapes? CanvasItems { get; }

    /// <summary>
    /// 获取指定对象的类型 ID。
    /// </summary>
    int ID { get; }

    /// <summary>
    /// 获取或设置一个值，指示表格中的形状是否在表格内部显示。
    /// </summary>
    int LayoutInCell { get; set; }

    /// <summary>
    /// 获取一个值，指示指定形状是否包含图表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasChart { get; }

    /// <summary>
    /// 获取表示文档形状集合中图表的 Chart 对象。
    /// </summary>
    IWordChart? Chart { get; }

    /// <summary>
    /// 获取或设置形状的相对水平位置。
    /// </summary>
    float LeftRelative { get; set; }

    /// <summary>
    /// 获取或设置形状的相对垂直位置。
    /// </summary>
    float TopRelative { get; set; }

    /// <summary>
    /// 获取或设置形状的相对宽度。
    /// </summary>
    float WidthRelative { get; set; }

    /// <summary>
    /// 获取或设置形状的相对高度百分比。
    /// </summary>
    float HeightRelative { get; set; }

    /// <summary>
    /// 获取或设置表示形状范围相对水平大小的常量。
    /// </summary>
    WdRelativeHorizontalSize RelativeHorizontalSize { get; set; }

    /// <summary>
    /// 获取或设置表示形状相对垂直大小的常量。
    /// </summary>
    WdRelativeVerticalSize RelativeVerticalSize { get; set; }

    /// <summary>
    /// 获取表示形状柔化边缘格式的 SoftEdgeFormat 对象。
    /// </summary>
    IWordSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取表示形状发光格式的 GlowFormat 对象。
    /// </summary>
    IWordGlowFormat? Glow { get; }

    /// <summary>
    /// 获取表示形状反射格式的 ReflectionFormat 对象。
    /// </summary>
    IWordReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取包含形状文本的 TextFrame2 对象。
    /// </summary>
    IOfficeTextFrame2? TextFrame2 { get; }

    /// <summary>
    /// 获取一个值，指示形状是否包含 SmartArt 图表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasSmartArt { get; }

    /// <summary>
    /// 获取提供与形状关联的 SmartArt 操作方式的 SmartArt 对象。
    /// </summary>
    IOfficeSmartArt? SmartArt { get; }

    /// <summary>
    /// 获取或设置指定形状的形状样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeStyleIndex ShapeStyle { get; set; }

    /// <summary>
    /// 获取或设置指定形状的背景样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取或设置包含指定形状标题的字符串。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 应用已使用 PickUp 方法复制的指定形状格式。
    /// </summary>
    void Apply();

    /// <summary>
    /// 删除图形。
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制指定图形，将新图形添加到图形集合中，并返回新图形。
    /// </summary>
    /// <returns>复制后的新图形对象。</returns>
    IWordShape? Duplicate();

    /// <summary>
    /// 水平或垂直翻转图形。
    /// </summary>
    /// <param name="flipCmd">翻转方向。</param>
    void Flip([ComNamespace("MsCore")] MsoFlipCmd flipCmd);

    /// <summary>
    /// 将图形水平移动指定距离。
    /// </summary>
    /// <param name="increment">图形水平移动的距离（以磅为单位）。正值向右移动，负值向左移动。</param>
    void IncrementLeft(float increment);

    /// <summary>
    /// 将图形围绕 Z 轴旋转指定度数。
    /// </summary>
    /// <param name="increment">图形水平旋转的度数。正值顺时针旋转，负值逆时针旋转。</param>
    void IncrementRotation(float increment);

    /// <summary>
    /// 将图形垂直移动指定距离。
    /// </summary>
    /// <param name="increment">图形垂直移动的距离（以磅为单位）。正值向下移动，负值向上移动。</param>
    void IncrementTop(float increment);

    /// <summary>
    /// 复制指定形状的格式。
    /// </summary>
    void PickUp();

    /// <summary>
    /// 按指定因子缩放图形高度。
    /// </summary>
    /// <param name="factor">缩放后高度与当前或原始高度的比率。例如，要使矩形高度增加50%，请指定1.5。</param>
    /// <param name="relativeToOriginalSize">True 表示相对于原始大小缩放，False 表示相对于当前大小缩放。仅当形状为图片或 OLE 对象时才能指定 True。</param>
    /// <param name="scale">缩放时保持位置不变的形状部分。</param>
    void ScaleHeight(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom scale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 按指定因子缩放图形宽度。
    /// </summary>
    /// <param name="factor">缩放后宽度与当前或原始宽度的比率。例如，要使矩形宽度增加50%，请指定1.5。</param>
    /// <param name="relativeToOriginalSize">True 表示相对于原始大小缩放，False 表示相对于当前大小缩放。仅当形状为图片或 OLE 对象时才能指定 True。</param>
    /// <param name="scale">缩放时保持位置不变的形状部分。</param>
    void ScaleWidth(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom scale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 选择图形。
    /// </summary>
    /// <param name="replace">如果添加形状，True 替换当前选择，False 将新形状添加到选择中。</param>
    void Select(bool? replace = null);

    /// <summary>
    /// 将指定形状的格式应用到该文档的默认形状。新形状从其默认形状继承许多属性。
    /// </summary>
    void SetShapesDefaultProperties();

    /// <summary>
    /// 取消指定形状中的任何已分组形状。
    /// </summary>
    /// <returns>取消分组后的形状范围。</returns>
    IWordShapeRange? Ungroup();

    /// <summary>
    /// 将指定形状移到其他形状的前面或后面（即更改形状在 Z 轴顺序中的位置）。
    /// </summary>
    /// <param name="zOrderCmd">指定将形状相对于其他形状移动到的位置。</param>
    void ZOrder([ComNamespace("MsCore")] MsoZOrderCmd zOrderCmd);

    /// <summary>
    /// 将文档绘图层中的指定形状转换为文本层中的内嵌形状。
    /// </summary>
    /// <returns>转换后的内嵌形状对象。</returns>
    IWordInlineShape? ConvertToInlineShape();

    /// <summary>
    /// 将指定形状转换为框架。
    /// </summary>
    /// <returns>转换后的框架对象。</returns>
    IWordFrame? ConvertToFrame();

    /// <summary>
    /// 从绘图画布的左侧裁剪指定百分比宽度。
    /// </summary>
    /// <param name="increment">裁剪后剩余的画布宽度百分比。例如，输入0.9表示从左侧裁剪10%宽度，输入0.1表示从左侧裁剪90%宽度。</param>
    void CanvasCropLeft(float increment);

    /// <summary>
    /// 从绘图画布的顶部裁剪指定百分比高度。
    /// </summary>
    /// <param name="increment">裁剪后剩余的画布高度百分比。例如，输入0.9表示从顶部裁剪10%高度，输入0.1表示从顶部裁剪90%高度。</param>
    void CanvasCropTop(float increment);

    /// <summary>
    /// 从绘图画布的右侧裁剪指定百分比宽度。
    /// </summary>
    /// <param name="increment">裁剪后剩余的画布宽度百分比。例如，输入0.9表示从右侧裁剪10%宽度，输入0.1表示从右侧裁剪90%宽度。</param>
    void CanvasCropRight(float increment);

    /// <summary>
    /// 从绘图画布的底部裁剪指定百分比高度。
    /// </summary>
    /// <param name="increment">裁剪后剩余的画布高度百分比。例如，输入0.9表示从底部裁剪10%高度，输入0.1表示从底部裁剪90%高度。</param>
    void CanvasCropBottom(float increment);


}