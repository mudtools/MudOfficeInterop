//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示幻灯片上的一个形状，提供对形状属性、格式和操作的全面访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointShape : IOfficeObject<IPowerPointShape, MsPowerPoint.Shape>, IDisposable
{
    /// <summary>
    /// 获取创建此形状的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取形状的父对象。
    /// </summary>
    /// <value>父对象，通常是幻灯片或组合形状。</value>
    object? Parent { get; }

    /// <summary>
    /// 应用形状的当前格式设置。
    /// </summary>
    void Apply();

    /// <summary>
    /// 删除形状。
    /// </summary>
    void Delete();

    /// <summary>
    /// 按指定方向翻转形状。
    /// </summary>
    /// <param name="flipCmd">翻转的方向。</param>
    void Flip([ComNamespace("MsCore")] MsoFlipCmd flipCmd);

    /// <summary>
    /// 将形状向左移动指定的距离。
    /// </summary>
    /// <param name="increment">向左移动的距离（磅）。</param>
    void IncrementLeft(float increment);

    /// <summary>
    /// 将形状旋转指定的角度。
    /// </summary>
    /// <param name="increment">旋转的角度（度）。</param>
    void IncrementRotation(float increment);

    /// <summary>
    /// 将形状向上移动指定的距离。
    /// </summary>
    /// <param name="increment">向上移动的距离（磅）。</param>
    void IncrementTop(float increment);

    /// <summary>
    /// 复制形状的格式设置。
    /// </summary>
    void PickUp();

    /// <summary>
    /// 重新路由连接符形状的连接点。
    /// </summary>
    void RerouteConnections();

    /// <summary>
    /// 按指定比例缩放形状的高度。
    /// </summary>
    /// <param name="factor">缩放比例因子（1.0 表示 100%）。</param>
    /// <param name="relativeToOriginalSize">指示是否相对于原始尺寸进行缩放。</param>
    /// <param name="scale">缩放基准点。</param>
    void ScaleHeight(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom scale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 按指定比例缩放形状的宽度。
    /// </summary>
    /// <param name="factor">缩放比例因子（1.0 表示 100%）。</param>
    /// <param name="relativeToOriginalSize">指示是否相对于原始尺寸进行缩放。</param>
    /// <param name="scale">缩放基准点。</param>
    void ScaleWidth(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom scale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 将当前形状的格式设置为默认格式。
    /// </summary>
    void SetShapesDefaultProperties();

    /// <summary>
    /// 取消组合形状，返回取消组合后的形状范围。
    /// </summary>
    /// <returns>包含取消组合后形状的形状范围。</returns>
    IPowerPointShapeRange? Ungroup();

    /// <summary>
    /// 更改形状在Z顺序中的位置。
    /// </summary>
    /// <param name="zOrderCmd">Z顺序操作命令。</param>
    void ZOrder([ComNamespace("MsCore")] MsoZOrderCmd zOrderCmd);

    /// <summary>
    /// 获取形状的调整值集合。
    /// </summary>
    /// <value>形状的调整值对象。</value>
    IPowerPointAdjustments? Adjustments { get; }

    /// <summary>
    /// 获取或设置自动形状的类型。
    /// </summary>
    /// <value>自动形状的类型。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoShapeType AutoShapeType { get; set; }

    /// <summary>
    /// 获取或设置形状在黑白模式下的显示方式。
    /// </summary>
    /// <value>黑白模式设置。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBlackWhiteMode BlackWhiteMode { get; set; }

    /// <summary>
    /// 获取形状的标注格式设置。
    /// </summary>
    /// <value>标注格式对象。</value>
    IPowerPointCalloutFormat? Callout { get; }

    /// <summary>
    /// 获取形状上的连接点数量。
    /// </summary>
    /// <value>连接点的总数。</value>
    int ConnectionSiteCount { get; }

    /// <summary>
    /// 获取一个值，指示形状是否为连接符。
    /// </summary>
    /// <value>如果形状是连接符，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Connector { get; }

    /// <summary>
    /// 获取形状的连接符格式设置。
    /// </summary>
    /// <value>连接符格式对象。</value>
    IPowerPointConnectorFormat? ConnectorFormat { get; }

    /// <summary>
    /// 获取形状的填充格式设置。
    /// </summary>
    /// <value>填充格式对象。</value>
    IPowerPointFillFormat? Fill { get; }

    /// <summary>
    /// 获取组合形状中的子形状集合。
    /// </summary>
    /// <value>组合形状中的子形状集合。</value>
    IPowerPointGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取或设置形状的高度（磅）。
    /// </summary>
    /// <value>形状的高度（磅）。</value>
    float Height { get; set; }

    /// <summary>
    /// 获取一个值，指示形状是否已水平翻转。
    /// </summary>
    /// <value>如果形状已水平翻转，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HorizontalFlip { get; }

    /// <summary>
    /// 获取或设置形状的左边缘相对于幻灯片左边缘的位置（磅）。
    /// </summary>
    /// <value>形状的左边缘位置（磅）。</value>
    float Left { get; set; }

    /// <summary>
    /// 获取形状的线条格式设置。
    /// </summary>
    /// <value>线条格式对象。</value>
    IPowerPointLineFormat? Line { get; }

    /// <summary>
    /// 获取或设置一个值，指示形状是否锁定纵横比。
    /// </summary>
    /// <value>如果形状锁定纵横比，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取或设置形状的名称。
    /// </summary>
    /// <value>形状的名称。</value>
    string? Name { get; set; }

    /// <summary>
    /// 获取形状的节点集合。
    /// </summary>
    /// <value>形状节点集合。</value>
    IPowerPointShapeNodes? Nodes { get; }

    /// <summary>
    /// 获取或设置形状的旋转角度（度）。
    /// </summary>
    /// <value>形状的旋转角度（度）。</value>
    float Rotation { get; set; }

    /// <summary>
    /// 获取形状的图片格式设置。
    /// </summary>
    /// <value>图片格式对象。</value>
    IPowerPointPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取形状的阴影格式设置。
    /// </summary>
    /// <value>阴影格式对象。</value>
    IPowerPointShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取形状的文本效果格式设置。
    /// </summary>
    /// <value>文本效果格式对象。</value>
    IPowerPointTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取形状的文本框架设置。
    /// </summary>
    /// <value>文本框架对象。</value>
    IPowerPointTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取形状的三维格式设置。
    /// </summary>
    /// <value>三维格式对象。</value>
    IPowerPointThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取或设置形状的上边缘相对于幻灯片上边缘的位置（磅）。
    /// </summary>
    /// <value>形状的上边缘位置（磅）。</value>
    float Top { get; set; }

    /// <summary>
    /// 获取形状的类型。
    /// </summary>
    /// <value>形状的类型。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取一个值，指示形状是否已垂直翻转。
    /// </summary>
    /// <value>如果形状已垂直翻转，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool VerticalFlip { get; }

    /// <summary>
    /// 获取形状的顶点坐标。
    /// </summary>
    /// <value>包含顶点坐标的数组。</value>
    object? Vertices { get; }

    /// <summary>
    /// 获取或设置一个值，指示形状是否可见。
    /// </summary>
    /// <value>如果形状可见，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度（磅）。
    /// </summary>
    /// <value>形状的宽度（磅）。</value>
    float Width { get; set; }

    /// <summary>
    /// 获取形状在Z顺序中的位置。
    /// </summary>
    /// <value>形状在Z顺序中的位置（1 表示最底层）。</value>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取形状的 OLE 格式设置。
    /// </summary>
    /// <value>OLE 格式对象。</value>
    IPowerPointOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取形状的链接格式设置。
    /// </summary>
    /// <value>链接格式对象。</value>
    IPowerPointLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取形状的占位符格式设置。
    /// </summary>
    /// <value>占位符格式对象。</value>
    IPowerPointPlaceholderFormat? PlaceholderFormat { get; }

    /// <summary>
    /// 获取形状的动画设置。
    /// </summary>
    /// <value>动画设置对象。</value>
    IPowerPointAnimationSettings? AnimationSettings { get; }

    /// <summary>
    /// 获取形状的动作设置。
    /// </summary>
    /// <value>动作设置对象。</value>
    IPowerPointActionSettings? ActionSettings { get; }

    /// <summary>
    /// 获取形状的标签集合。
    /// </summary>
    /// <value>标签集合对象。</value>
    IPowerPointTags? Tags { get; }

    /// <summary>
    /// 剪切形状。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制形状。
    /// </summary>
    void Copy();

    /// <summary>
    /// 选择形状。
    /// </summary>
    /// <param name="replace">指示是否替换当前选择。</param>
    void Select([ConvertTriState] bool replace = true);

    /// <summary>
    /// 复制形状并返回新形状。
    /// </summary>
    /// <returns>新复制的形状范围。</returns>
    IPowerPointShapeRange? Duplicate();

    /// <summary>
    /// 获取形状的媒体类型。
    /// </summary>
    /// <value>媒体类型。</value>
    PpMediaType MediaType { get; }

    /// <summary>
    /// 获取一个值，指示形状是否具有文本框架。
    /// </summary>
    /// <value>如果形状具有文本框架，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasTextFrame { get; }

    /// <summary>
    /// 获取形状的声音格式设置。
    /// </summary>
    /// <value>声音格式对象。</value>
    IPowerPointSoundFormat? SoundFormat { get; }

    /// <summary>
    /// 获取形状的脚本对象。
    /// </summary>
    /// <value>脚本对象。</value>
    IOfficeScript? Script { get; }

    /// <summary>
    /// 获取或设置形状的替代文本。
    /// </summary>
    /// <value>形状的替代文本。</value>
    string? AlternativeText { get; set; }

    /// <summary>
    /// 获取一个值，指示形状是否包含表格。
    /// </summary>
    /// <value>如果形状包含表格，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasTable { get; }

    /// <summary>
    /// 获取形状中的表格对象。
    /// </summary>
    /// <value>表格对象。</value>
    IPowerPointTable? Table { get; }

    /// <summary>
    /// 将形状导出为图像文件。
    /// </summary>
    /// <param name="pathName">导出文件的路径和名称。</param>
    /// <param name="filter">导出文件的格式。</param>
    /// <param name="scaleWidth">导出图像的缩放宽度（像素）。</param>
    /// <param name="scaleHeight">导出图像的缩放高度（像素）。</param>
    /// <param name="exportMode">导出模式。</param>
    void Export(string pathName, PpShapeFormat filter, int scaleWidth = 0, int scaleHeight = 0, PpExportMode exportMode = PpExportMode.ppRelativeToSlide);

    /// <summary>
    /// 获取一个值，指示形状是否包含图示。
    /// </summary>
    /// <value>如果形状包含图示，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasDiagram { get; }

    /// <summary>
    /// 获取形状中的图示对象。
    /// </summary>
    /// <value>图示对象。</value>
    IPowerPointDiagram? Diagram { get; }

    /// <summary>
    /// 获取一个值，指示形状是否包含图示节点。
    /// </summary>
    /// <value>如果形状包含图示节点，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasDiagramNode { get; }

    /// <summary>
    /// 获取形状中的图示节点对象。
    /// </summary>
    /// <value>图示节点对象。</value>
    IPowerPointDiagramNode? DiagramNode { get; }

    /// <summary>
    /// 获取一个值，指示形状是否为子形状。
    /// </summary>
    /// <value>如果形状是子形状，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Child { get; }

    /// <summary>
    /// 获取形状的父组合形状。
    /// </summary>
    /// <value>父组合形状。</value>
    IPowerPointShape? ParentGroup { get; }

    /// <summary>
    /// 获取形状的画布形状集合。
    /// </summary>
    /// <value>画布形状集合。</value>
    IPowerPointCanvasShapes? CanvasItems { get; }

    /// <summary>
    /// 获取形状的标识符。
    /// </summary>
    /// <value>形状的唯一标识符。</value>
    int Id { get; }

    /// <summary>
    /// 从画布左侧裁剪形状。
    /// </summary>
    /// <param name="increment">裁剪的距离（磅）。</param>
    void CanvasCropLeft(float increment);

    /// <summary>
    /// 从画布顶部裁剪形状。
    /// </summary>
    /// <param name="increment">裁剪的距离（磅）。</param>
    void CanvasCropTop(float increment);

    /// <summary>
    /// 从画布右侧裁剪形状。
    /// </summary>
    /// <param name="increment">裁剪的距离（磅）。</param>
    void CanvasCropRight(float increment);

    /// <summary>
    /// 从画布底部裁剪形状。
    /// </summary>
    /// <param name="increment">裁剪的距离（磅）。</param>
    void CanvasCropBottom(float increment);

    /// <summary>
    /// 设置形状的 RTF 格式文本。
    /// </summary>
    /// <value>RTF 格式文本。</value>
    string? RTF { set; }

    /// <summary>
    /// 获取形状的自定义数据。
    /// </summary>
    /// <value>自定义数据对象。</value>
    IPowerPointCustomerData? CustomerData { get; }

    /// <summary>
    /// 获取形状的文本框架2设置。
    /// </summary>
    /// <value>文本框架2对象。</value>
    IPowerPointTextFrame2? TextFrame2 { get; }

    /// <summary>
    /// 获取一个值，指示形状是否包含图表。
    /// </summary>
    /// <value>如果形状包含图表，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasChart { get; }

    /// <summary>
    /// 获取或设置形状的样式。
    /// </summary>
    /// <value>形状样式索引。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeStyleIndex ShapeStyle { get; set; }

    /// <summary>
    /// 获取或设置形状的背景样式。
    /// </summary>
    /// <value>背景样式索引。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取形状的柔化边缘效果设置。
    /// </summary>
    /// <value>柔化边缘格式对象。</value>
    IOfficeSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取形状的发光效果设置。
    /// </summary>
    /// <value>发光格式对象。</value>
    IOfficeGlowFormat? Glow { get; }

    /// <summary>
    /// 获取形状的反射效果设置。
    /// </summary>
    /// <value>反射格式对象。</value>
    IOfficeReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取形状中的图表对象。
    /// </summary>
    /// <value>图表对象。</value>
    //IPowerPointChart? Chart { get; }

    /// <summary>
    /// 获取一个值，指示形状是否包含 SmartArt 图形。
    /// </summary>
    /// <value>如果形状包含 SmartArt 图形，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasSmartArt { get; }

    /// <summary>
    /// 获取形状中的 SmartArt 图形对象。
    /// </summary>
    /// <value>SmartArt 图形对象。</value>
    IOfficeSmartArt? SmartArt { get; }

    /// <summary>
    /// 将形状中的文本转换为 SmartArt 图形。
    /// </summary>
    /// <param name="layout">要应用的 SmartArt 布局。</param>
    void ConvertTextToSmartArt(IOfficeSmartArtLayout layout);

    /// <summary>
    /// 获取或设置形状的标题。
    /// </summary>
    /// <value>形状的标题文本。</value>
    string? Title { get; set; }

    /// <summary>
    /// 获取形状的媒体格式设置。
    /// </summary>
    /// <value>媒体格式对象。</value>
    IPowerPointMediaFormat? MediaFormat { get; }

    /// <summary>
    /// 复制形状的动画效果。
    /// </summary>
    void PickupAnimation();

    /// <summary>
    /// 应用复制的动画效果到形状。
    /// </summary>
    void ApplyAnimation();

    /// <summary>
    /// 升级形状中的媒体格式。
    /// </summary>
    void UpgradeMedia();
}