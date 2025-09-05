namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定Word文档中内嵌形状的类型
/// </summary>
public enum WdInlineShapeType
{
    /// <summary>
    /// 嵌入式OLE对象
    /// </summary>
    wdInlineShapeEmbeddedOLEObject = 1,
    /// <summary>
    /// 链接式OLE对象
    /// </summary>
    wdInlineShapeLinkedOLEObject,
    /// <summary>
    /// 图片
    /// </summary>
    wdInlineShapePicture,
    /// <summary>
    /// 链接式图片
    /// </summary>
    wdInlineShapeLinkedPicture,
    /// <summary>
    /// OLE控件对象
    /// </summary>
    wdInlineShapeOLEControlObject,
    /// <summary>
    /// 水平线
    /// </summary>
    wdInlineShapeHorizontalLine,
    /// <summary>
    /// 带水平线的图片
    /// </summary>
    wdInlineShapePictureHorizontalLine,
    /// <summary>
    /// 带水平线的链接式图片
    /// </summary>
    wdInlineShapeLinkedPictureHorizontalLine,
    /// <summary>
    /// 项目符号图片
    /// </summary>
    wdInlineShapePictureBullet,
    /// <summary>
    /// 脚本锚点
    /// </summary>
    wdInlineShapeScriptAnchor,
    /// <summary>
    /// OWS锚点
    /// </summary>
    wdInlineShapeOWSAnchor,
    /// <summary>
    /// 图表
    /// </summary>
    wdInlineShapeChart,
    /// <summary>
    /// 图解
    /// </summary>
    wdInlineShapeDiagram,
    /// <summary>
    /// 锁定画布
    /// </summary>
    wdInlineShapeLockedCanvas,
    /// <summary>
    /// SmartArt图形
    /// </summary>
    wdInlineShapeSmartArt,
    /// <summary>
    /// 网络视频
    /// </summary>
    wdInlineShapeWebVideo
}