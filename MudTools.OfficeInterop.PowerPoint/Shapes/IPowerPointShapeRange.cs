//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 中形状的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointShapeRange : IDisposable, IEnumerable<IPowerPointShape?>
{
    /// <summary>
    /// 获取创建此形状集合的应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此形状集合的应用程序的创建者代码。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取此形状集合的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 应用已通过 <see cref="PickUp"/> 方法复制的格式。
    /// </summary>
    void Apply();

    /// <summary>
    /// 删除形状集合。
    /// </summary>
    void Delete();

    /// <summary>
    /// 沿指定方向翻转形状集合。
    /// </summary>
    /// <param name="flipCmd">指定翻转的方向。</param>
    void Flip([ComNamespace("MsCore")] MsoFlipCmd flipCmd);

    /// <summary>
    /// 将形状集合水平移动指定的量。
    /// </summary>
    /// <param name="increment">水平移动的距离（以磅为单位）。正数向右移动，负数向左移动。</param>
    void IncrementLeft(float increment);

    /// <summary>
    /// 将形状集合按指定角度旋转。
    /// </summary>
    /// <param name="increment">旋转的角度（以度为单位）。正数顺时针旋转，负数逆时针旋转。</param>
    void IncrementRotation(float increment);

    /// <summary>
    /// 将形状集合垂直移动指定的量。
    /// </summary>
    /// <param name="increment">垂直移动的距离（以磅为单位）。正数向下移动，负数向上移动。</param>
    void IncrementTop(float increment);

    /// <summary>
    /// 复制形状集合的格式，以便使用 <see cref="Apply"/> 方法将其应用于其他形状。
    /// </summary>
    void PickUp();

    /// <summary>
    /// 重排形状集合中连接符的形状，以便它们采用最短路径。
    /// </summary>
    void RerouteConnections();

    /// <summary>
    /// 按指定因子缩放形状集合的高度。
    /// </summary>
    /// <param name="factor">缩放因子（例如，1.5 表示 150%）。</param>
    /// <param name="relativeToOriginalSize">指示缩放是否相对于原始大小。</param>
    /// <param name="fScale">指定从哪一侧开始缩放。</param>
    void ScaleHeight(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 按指定因子缩放形状集合的宽度。
    /// </summary>
    /// <param name="factor">缩放因子（例如，1.5 表示 150%）。</param>
    /// <param name="relativeToOriginalSize">指示缩放是否相对于原始大小。</param>
    /// <param name="fScale">指定从哪一侧开始缩放。</param>
    void ScaleWidth(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom fScale = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 将形状集合的格式设置为新形状的默认格式。
    /// </summary>
    void SetShapesDefaultProperties();

    /// <summary>
    /// 取消形状集合的分组，返回取消分组后的形状范围。
    /// </summary>
    /// <returns>表示取消分组后形状的 <see cref="IPowerPointShapeRange"/>。</returns>
    IPowerPointShapeRange? Ungroup();

    /// <summary>
    /// 将形状集合置于其他形状的前面或后面。
    /// </summary>
    /// <param name="zOrderCmd">指定排序命令。</param>
    void ZOrder([ComNamespace("MsCore")] MsoZOrderCmd zOrderCmd);

    /// <summary>
    /// 获取形状集合的调整值。
    /// </summary>
    IPowerPointAdjustments? Adjustments { get; }

    /// <summary>
    /// 获取或设置形状集合的自动形状类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoShapeType AutoShapeType { get; set; }

    /// <summary>
    /// 获取或设置形状集合以黑白模式显示时的外观。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBlackWhiteMode BlackWhiteMode { get; set; }

    /// <summary>
    /// 获取形状集合的标注格式。
    /// </summary>
    IPowerPointCalloutFormat? Callout { get; }

    /// <summary>
    /// 获取形状集合中的连接点数量。
    /// </summary>
    int ConnectionSiteCount { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否为连接符。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Connector { get; }

    /// <summary>
    /// 获取形状集合的连接符格式。
    /// </summary>
    IPowerPointConnectorFormat? ConnectorFormat { get; }

    /// <summary>
    /// 获取形状集合的填充格式。
    /// </summary>
    IPowerPointFillFormat? Fill { get; }

    /// <summary>
    /// 获取构成形状集合中分组形状的单个形状。
    /// </summary>
    IPowerPointGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取或设置形状集合的高度（以磅为单位）。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取一个值，指示形状集合是否已水平翻转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HorizontalFlip { get; }

    /// <summary>
    /// 获取或设置形状集合左边缘到幻灯片左边缘的距离（以磅为单位）。
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取形状集合的线条格式。
    /// </summary>
    IPowerPointLineFormat? Line { get; }

    /// <summary>
    /// 获取或设置一个值，指示形状集合的宽高比是否锁定。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取或设置形状集合的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状集合的节点。
    /// </summary>
    IPowerPointShapeNodes? Nodes { get; }

    /// <summary>
    /// 获取或设置形状集合的旋转角度（以度为单位）。
    /// </summary>
    float Rotation { get; set; }

    /// <summary>
    /// 获取形状集合的图片格式。
    /// </summary>
    IPowerPointPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取形状集合的阴影格式。
    /// </summary>
    IPowerPointShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取形状集合的文本效果格式。
    /// </summary>
    IPowerPointTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取形状集合的文本框。
    /// </summary>
    IPowerPointTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取形状集合的三维格式。
    /// </summary>
    IPowerPointThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取或设置形状集合上边缘到幻灯片上边缘的距离（以磅为单位）。
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取形状集合的类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否已垂直翻转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool VerticalFlip { get; }

    /// <summary>
    /// 获取形状集合的顶点坐标。
    /// </summary>
    object Vertices { get; }

    /// <summary>
    /// 获取或设置一个值，指示形状集合是否可见。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置形状集合的宽度（以磅为单位）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取形状集合在 z 轴顺序中的位置。
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取形状集合的 OLE 格式。
    /// </summary>
    IPowerPointOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取形状集合的链接格式。
    /// </summary>
    IPowerPointLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取形状集合的占位符格式。
    /// </summary>
    IPowerPointPlaceholderFormat? PlaceholderFormat { get; }

    /// <summary>
    /// 获取形状集合的动画设置。
    /// </summary>
    IPowerPointAnimationSettings? AnimationSettings { get; }

    /// <summary>
    /// 获取形状集合的动作设置。
    /// </summary>
    IPowerPointActionSettings? ActionSettings { get; }

    /// <summary>
    /// 获取形状集合的标签。
    /// </summary>
    IPowerPointTags? Tags { get; }

    /// <summary>
    /// 将形状集合剪切到剪贴板。
    /// </summary>
    void Cut();

    /// <summary>
    /// 将形状集合复制到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 选择形状集合。
    /// </summary>
    /// <param name="replace">指示是否替换当前选择。</param>
    void Select([ConvertTriState] bool replace = true);

    /// <summary>
    /// 复制形状集合并返回对副本的引用。
    /// </summary>
    /// <returns>表示复制后形状的 <see cref="IPowerPointShapeRange"/>。</returns>
    IPowerPointShapeRange? Duplicate();

    /// <summary>
    /// 获取形状集合的媒体类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsPowerPoint")]
    PpMediaType MediaType { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否具有文本框。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasTextFrame { get; }

    /// <summary>
    /// 获取形状集合的声音格式。
    /// </summary>
    IPowerPointSoundFormat? SoundFormat { get; }

    /// <summary>
    /// 获取形状集合中指定索引处的形状。
    /// </summary>
    /// <param name="index">形状的索引或名称。</param>
    /// <returns>指定索引处的 <see cref="IPowerPointShape"/>。</returns>
    IPowerPointShape? this[object index] { get; }

    /// <summary>
    /// 获取形状集合中形状的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 将形状集合分组并返回分组的形状。
    /// </summary>
    /// <returns>表示分组后形状的 <see cref="IPowerPointShape"/>。</returns>
    IPowerPointShape? Group();

    /// <summary>
    /// 重新组合之前已分组的形状集合并返回重新组合的形状。
    /// </summary>
    /// <returns>表示重新组合后形状的 <see cref="IPowerPointShape"/>。</returns>
    IPowerPointShape? Regroup();

    /// <summary>
    /// 对齐形状集合中的形状。
    /// </summary>
    /// <param name="alignCmd">指定对齐方式。</param>
    /// <param name="relativeTo">指示对齐是相对于幻灯片还是相对于彼此。</param>
    void Align([ComNamespace("MsCore")] MsoAlignCmd alignCmd, [ConvertTriState] bool relativeTo);

    /// <summary>
    /// 在形状集合中均匀分布形状。
    /// </summary>
    /// <param name="distributeCmd">指定分布方式。</param>
    /// <param name="relativeTo">指示分布是相对于幻灯片还是相对于彼此。</param>
    void Distribute([ComNamespace("MsCore")] MsoDistributeCmd distributeCmd, [ConvertTriState] bool relativeTo);

    /// <summary>
    /// 获取形状集合的脚本。
    /// </summary>
    IOfficeScript? Script { get; }

    /// <summary>
    /// 获取或设置形状集合的替代文本。
    /// </summary>
    string AlternativeText { get; set; }

    /// <summary>
    /// 获取一个值，指示形状集合是否具有表格。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasTable { get; }

    /// <summary>
    /// 获取形状集合中的表格。
    /// </summary>
    IPowerPointTable? Table { get; }

    /// <summary>
    /// 将形状集合导出为图像文件。
    /// </summary>
    /// <param name="pathName">导出文件的路径和名称。</param>
    /// <param name="filter">导出的图像格式。</param>
    /// <param name="scaleWidth">缩放后的宽度（以像素为单位）。</param>
    /// <param name="scaleHeight">缩放后的高度（以像素为单位）。</param>
    /// <param name="exportMode">指定导出模式。</param>
    void Export(string pathName, PpShapeFormat filter, int scaleWidth = 0, int scaleHeight = 0, PpExportMode exportMode = PpExportMode.ppRelativeToSlide);

    /// <summary>
    /// 获取一个值，指示形状集合是否具有图表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasDiagram { get; }

    /// <summary>
    /// 获取形状集合中的图表。
    /// </summary>
    IPowerPointDiagram? Diagram { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否具有图表节点。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasDiagramNode { get; }

    /// <summary>
    /// 获取形状集合中的图表节点。
    /// </summary>
    IPowerPointDiagramNode? DiagramNode { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否为子形状。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Child { get; }

    /// <summary>
    /// 获取形状集合的父组形状。
    /// </summary>
    IPowerPointShape? ParentGroup { get; }

    /// <summary>
    /// 获取形状集合的画布形状。
    /// </summary>
    IPowerPointCanvasShapes? CanvasItems { get; }

    /// <summary>
    /// 获取形状集合的标识符。
    /// </summary>
    int Id { get; }

    /// <summary>
    /// 从画布形状的左侧裁剪指定的量。
    /// </summary>
    /// <param name="increment">裁剪的量（以磅为单位）。</param>
    void CanvasCropLeft(float increment);

    /// <summary>
    /// 从画布形状的顶部裁剪指定的量。
    /// </summary>
    /// <param name="increment">裁剪的量（以磅为单位）。</param>
    void CanvasCropTop(float increment);

    /// <summary>
    /// 从画布形状的右侧裁剪指定的量。
    /// </summary>
    /// <param name="increment">裁剪的量（以磅为单位）。</param>
    void CanvasCropRight(float increment);

    /// <summary>
    /// 从画布形状的底部裁剪指定的量。
    /// </summary>
    /// <param name="increment">裁剪的量（以磅为单位）。</param>
    void CanvasCropBottom(float increment);

    /// <summary>
    /// 设置形状集合的 RTF 格式文本。
    /// </summary>
    string RTF { set; }

    /// <summary>
    /// 获取形状集合的客户数据。
    /// </summary>
    IPowerPointCustomerData? CustomerData { get; }

    /// <summary>
    /// 获取形状集合的文本框2。
    /// </summary>
    IPowerPointTextFrame2? TextFrame2 { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否具有图表。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasChart { get; }

    /// <summary>
    /// 获取或设置形状集合的形状样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeStyleIndex ShapeStyle { get; set; }

    /// <summary>
    /// 获取或设置形状集合的背景样式。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取形状集合的柔化边缘格式。
    /// </summary>
    IOfficeSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取形状集合的发光格式。
    /// </summary>
    IOfficeGlowFormat? Glow { get; }

    /// <summary>
    /// 获取形状集合的反射格式。
    /// </summary>
    IOfficeReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取形状集合中的图表对象。
    /// </summary>
    //IPowerPointChart? Chart { get; }

    /// <summary>
    /// 获取一个值，指示形状集合是否具有 SmartArt 图形。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasSmartArt { get; }

    /// <summary>
    /// 获取形状集合中的 SmartArt 图形。
    /// </summary>
    IOfficeSmartArt? SmartArt { get; }

    /// <summary>
    /// 将形状集合中的文本转换为 SmartArt 图形。
    /// </summary>
    /// <param name="layout">要应用的 SmartArt 布局。</param>
    void ConvertTextToSmartArt(IOfficeSmartArtLayout layout);

    /// <summary>
    /// 获取或设置形状集合的标题。
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取形状集合的媒体格式。
    /// </summary>
    IPowerPointMediaFormat? MediaFormat { get; }

    /// <summary>
    /// 复制动画以便使用 <see cref="ApplyAnimation"/> 方法将其应用于其他形状。
    /// </summary>
    void PickupAnimation();

    /// <summary>
    /// 将已通过 <see cref="PickupAnimation"/> 方法复制的动画应用于形状集合。
    /// </summary>
    void ApplyAnimation();

    /// <summary>
    /// 升级形状集合中的媒体以支持最新的媒体格式。
    /// </summary>
    void UpgradeMedia();

    /// <summary>
    /// 将形状集合与另一个形状合并以创建新形状。
    /// </summary>
    /// <param name="mergeCmd">指定合并操作。</param>
    /// <param name="primaryShape">用作合并基准的主要形状。</param>
    void MergeShapes([ComNamespace("MsCore")] MsoMergeCmd mergeCmd, IPowerPointShape primaryShape = null);
}