//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel Shape 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.Shape 的安全访问和操作
/// </summary>
public interface IExcelShape : IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取形状的 OLE 格式设置属性
    /// 对应 Shape.OLEFormat 属性，提供对嵌入的 OLE 对象格式设置的访问
    /// </summary>
    IExcelOLEFormat? OLEFormat { get; }

    /// <summary>
    /// 获取组合形状中单个子形状的集合
    /// 对应 Shape.GroupItems 属性，仅当形状为组合形状时可用
    /// </summary>
    IExcelGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取形状的连接符格式设置属性
    /// 对应 Shape.ConnectorFormat 属性，用于控制连接符的类型、起始/终止连接对象及连接点
    /// </summary>
    IExcelConnectorFormat? ConnectorFormat { get; }

    /// <summary>
    /// 获取自由形状中所有路径节点的集合
    /// 对应 Shape.Nodes 属性，支持遍历、索引访问和节点操作
    /// </summary>
    IExcelShapeNodes? ShapeNodes { get; }

    /// <summary>
    /// 获取形状的链接格式设置属性
    /// 对应 Shape.LinkFormat 属性，用于管理链接源、更新方式、断开链接等操作
    /// </summary>
    IExcelLinkFormat? LinkFormat { get; }

    /// <summary>
    /// 获取形状的控件格式设置属性
    /// 对应 Shape.ControlFormat 属性，用于管理表单控件的列表项、当前值、范围、多选等属性
    /// </summary>
    IExcelControlFormat? ControlFormat { get; }

    /// <summary>
    /// 获取形状的柔化边缘格式设置属性
    /// 对应 Office.Interop 中的 SoftEdgeFormat 对象
    /// </summary>
    IOfficeSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取形状的发光格式设置属性
    /// 对应 Office.Interop 中的 GlowFormat 对象
    /// </summary>
    IOfficeGlowFormat? Glow { get; }

    /// <summary>
    /// 获取形状的父级组合形状
    /// 当前形状是组合形状的一部分时，返回其父级组合形状
    /// </summary>
    IExcelShape? ParentGroup { get; }

    /// <summary>
    /// 获取形状中的 SmartArt 对象
    /// 当形状包含 SmartArt 图形时可用
    /// </summary>
    IOfficeSmartArt? SmartArt { get; }

    /// <summary>
    /// 获取一个值，该值指示形状是否包含图表
    /// </summary>
    bool HasChart { get; }

    /// <summary>
    /// 获取或设置形状样式
    /// 对应 Office.Core 中的 MsoShapeStyleIndex 枚举值
    /// </summary>
    MsoShapeStyleIndex ShapeStyle { get; set; }

    /// <summary>
    /// 获取或设置形状背景样式
    /// 对应 Office.Core 中的 MsoBackgroundStyleIndex 枚举值
    /// </summary>
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取或设置形状的标题
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取或设置形状的名称
    /// 对应 Shape.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状的类型
    /// 对应 Shape.Type 属性
    /// </summary>
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取形状的ID
    /// 对应 Shape.ID 属性
    /// </summary>
    int ID { get; }

    /// <summary>
    /// 获取形状的父对象
    /// 对应 Shape.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置形状的定位方式
    /// 对应 Shape.Placement 属性
    /// </summary>
    XlPlacement Placement { get; set; }

    /// <summary>
    /// 获取形状是否为连接符
    /// 对应 Shape.Connector 属性，用于判断形状是否为连接符类型
    /// </summary>
    bool Connector { get; }

    /// <summary>
    /// 获取或设置形状的宽高比锁定状态
    /// 对应 Shape.LockAspectRatio 属性，当设置为 true 时，调整形状大小会保持原始宽高比
    /// </summary>
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取形状是否已水平翻转
    /// 对应 Shape.HorizontalFlip 属性，用于判断形状是否经过水平翻转
    /// </summary>
    bool HorizontalFlip { get; }

    /// <summary>
    /// 获取或设置自选图形的类型
    /// 对应 Shape.AutoShapeType 属性，用于指定自选图形的具体类型（如矩形、圆形等）
    /// </summary>
    MsoAutoShapeType AutoShapeType { get; set; }

    /// <summary>
    /// 获取或设置形状在黑白模式下的显示方式
    /// 对应 Shape.BlackWhiteMode 属性，用于控制形状在黑白视图中的显示效果
    /// </summary>
    MsoBlackWhiteMode BlackWhiteMode { get; set; }

    /// <summary>
    /// 获取表单控件的类型
    /// 对应 Shape.FormControlType 属性，用于确定表单控件的具体类型（如按钮、复选框等）
    /// </summary>
    XlFormControl FormControlType { get; }

    #endregion

    #region 位置和大小
    /// <summary>
    /// 获取或设置形状的左边距（以磅为单位）
    /// 对应 Shape.Left 属性
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取或设置形状的顶边距（以磅为单位）
    /// 对应 Shape.Top 属性
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取或设置形状的宽度（以磅为单位）
    /// 对应 Shape.Width 属性
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取或设置形状的高度（以磅为单位）
    /// 对应 Shape.Height 属性
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取或设置形状的旋转角度（以度为单位）
    /// 对应 Shape.Rotation 属性
    /// </summary>
    float Rotation { get; set; }

    #endregion

    #region 可见性和状态

    /// <summary>
    /// 获取或设置形状是否可见
    /// 对应 Shape.Visible 属性
    /// </summary>
    bool Visible { get; set; }

    /// <summary>
    /// 获取或设置形状是否锁定
    /// 对应 Shape.Locked 属性
    /// </summary>
    bool Locked { get; set; }
    #endregion

    #region 格式设置

    /// <summary>
    /// 获取形状在 Z 轴上的堆叠顺序位置
    /// 对应 Shape.ZOrderPosition 属性，返回形状在当前工作表中相对于其他形状的堆叠顺序
    /// </summary>
    int ZOrderPosition { get; }

    /// <summary>
    /// 获取形状的标注线格式设置属性
    /// 对应 Shape.Callout 属性，用于控制标注线的类型、角度、长度等属性
    /// </summary>
    IExcelCalloutFormat? Callout { get; }

    /// <summary>
    /// 获取形状的图片格式设置属性
    /// 对应 Shape.PictureFormat 属性，用于控制图片的亮度、对比度、透明度、裁剪等属性
    /// </summary>
    IExcelPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取形状的文本特效格式设置属性
    /// 对应 Shape.TextEffect 属性，用于控制文本特效的字体、大小、样式等属性
    /// </summary>
    IExcelTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取与形状关联的超链接对象
    /// 对应 Shape.Hyperlink 属性，用于访问和管理形状上的超链接设置
    /// </summary>
    IExcelHyperlink? Hyperlink { get; }

    /// <summary>
    /// 获取形状的填充格式对象
    /// 对应 Shape.Fill 属性，用于设置形状的填充颜色、渐变、图案等外观特性
    /// </summary>
    IExcelFillFormat Fill { get; }

    /// <summary>
    /// 获取形状的线条格式对象
    /// 对应 Shape.Line 属性，用于设置形状轮廓线的颜色、粗细、样式等外观特性
    /// </summary>
    IExcelLineFormat Line { get; }

    /// <summary>
    /// 获取形状的文本框架对象
    /// 对应 Shape.TextFrame 属性，用于控制文本的边距、方向、自动调整等布局属性
    /// </summary>
    IExcelTextFrame TextFrame { get; }

    /// <summary>
    /// 获取形状的阴影格式对象
    /// 对应 Shape.Shadow 属性，用于设置形状阴影的颜色、偏移量、模糊度等视觉效果
    /// </summary>
    IExcelShadowFormat Shadow { get; }

    /// <summary>
    /// 获取形状的三维格式对象
    /// 对应 Shape.ThreeD 属性，用于设置形状的深度、透视、表面材质等三维效果
    /// </summary>
    IExcelThreeDFormat ThreeD { get; }
    #endregion

    #region 文本属性

    /// <summary>
    /// 获取或设置形状中的文本内容
    /// 对应 Shape.TextFrame.Characters.Text 属性
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置形状中文本的自动调整大小
    /// 对应 Shape.TextFrame.AutoSize 属性
    /// </summary>
    bool AutoSize { get; set; }

    /// <summary>
    /// 获取或设置形状中文本的水平对齐方式
    /// 对应 Shape.TextFrame.HorizontalAlignment 属性
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置形状中文本的垂直对齐方式
    /// 对应 Shape.TextFrame.VerticalAlignment 属性
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    #endregion

    #region 操作方法

    /// <summary>
    /// 选择形状
    /// 对应 Shape.Select 方法
    /// </summary>
    /// <param name="replace">true表示替换当前选择，false表示添加到当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 复制形状
    /// 对应 Shape.Copy 方法
    /// </summary>
    void Copy();

    /// <summary>
    /// 复制形状
    /// 对应 Shape.Copy 方法
    /// </summary>
    void CopyPicture(XlPictureAppearance? Appearance, XlCopyPictureFormat? Format);

    /// <summary>
    /// 剪切形状
    /// 对应 Shape.Cut 方法
    /// </summary>
    void Cut();

    /// <summary>
    /// 删除形状
    /// 对应 Shape.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 高度缩放
    /// </summary>
    /// <param name="Factor"></param>
    /// <param name="RelativeToOriginalSize">是否相对于原始大小</param>
    /// <param name="Scale">缩放比例</param>
    void ScaleHeight(float Factor, bool RelativeToOriginalSize, float Scale);

    /// <summary>
    /// 宽度缩放
    /// </summary>
    /// <param name="Factor"></param>
    /// <param name="RelativeToOriginalSize">是否相对于原始大小</param>
    /// <param name="Scale">缩放比例</param>
    void ScaleWidth(float Factor, bool RelativeToOriginalSize, float Scale);

    /// <summary>
    /// 调整形状大小
    /// 对应 Shape.ScaleWidth 和 Shape.ScaleHeight 方法
    /// </summary>
    /// <param name="widthScale">宽度缩放比例</param>
    /// <param name="heightScale">高度缩放比例</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小</param>
    void Scale(float widthScale, float heightScale, bool relativeToOriginalSize = false);

    /// <summary>
    /// 移动形状
    /// 对应 Shape.IncrementLeft 和 Shape.IncrementTop 方法
    /// </summary>
    /// <param name="leftIncrement">左边距增量</param>
    /// <param name="topIncrement">顶边距增量</param>
    void Move(float leftIncrement, float topIncrement);

    /// <summary>
    /// 旋转形状
    /// 对应 Shape.IncrementRotation 方法
    /// </summary>
    /// <param name="rotationIncrement">旋转角度增量（度）</param>
    void Rotate(float rotationIncrement);

    /// <summary>
    /// 设置形状的堆叠顺序（Z轴顺序）
    /// 对应 Shape.ZOrder 方法
    /// </summary>
    /// <param name="orderCmd">Z轴顺序命令，指定如何重新排列对象的堆叠顺序</param>
    void ZOrder(MsoZOrderCmd orderCmd);

    /// <summary>
    /// 按指定增量调整形状的水平位置
    /// 对应 Shape.IncrementLeft 方法
    /// </summary>
    /// <param name="Increment">水平位置增量，以磅为单位</param>
    void IncrementLeft(float Increment);

    /// <summary>
    /// 按指定增量调整形状的垂直位置
    /// 对应 Shape.IncrementTop 方法
    /// </summary>
    /// <param name="Increment">垂直位置增量，以磅为单位</param>
    void IncrementTop(float Increment);

    /// <summary>
    /// 绕水平或垂直轴翻转形状
    /// 对应 Shape.Flip 方法
    /// </summary>
    /// <param name="FlipCmd">翻转方向命令，指定是水平翻转还是垂直翻转</param>
    void Flip(MsoFlipCmd FlipCmd);

    /// <summary>
    /// 重新路由任何连接符附加到该形状的连接点
    /// 对应 Shape.RerouteConnections 方法
    /// </summary>
    void RerouteConnections();

    /// <summary>
    /// 裁剪画布的左侧
    /// 对应 Shape.CanvasCropLeft 方法
    /// </summary>
    /// <param name="Increment">裁剪增量，以磅为单位</param>
    void CanvasCropLeft(float Increment);

    /// <summary>
    /// 裁剪画布的顶部
    /// 对应 Shape.CanvasCropTop 方法
    /// </summary>
    /// <param name="Increment">裁剪增量，以磅为单位</param>
    void CanvasCropTop(float Increment);

    /// <summary>
    /// 裁剪画布的右侧
    /// 对应 Shape.CanvasCropRight 方法
    /// </summary>
    /// <param name="Increment">裁剪增量，以磅为单位</param>
    void CanvasCropRight(float Increment);

    /// <summary>
    /// 裁剪画布的底部
    /// 对应 Shape.CanvasCropBottom 方法
    /// </summary>
    /// <param name="Increment">裁剪增量，以磅为单位</param>
    void CanvasCropBottom(float Increment);

    /// <summary>
    /// 将形状置于最前面
    /// 对应 Shape.ZOrder 方法
    /// </summary>
    void BringToFront();

    /// <summary>
    /// 将形状置于最后面
    /// 对应 Shape.ZOrder 方法
    /// </summary>
    void SendToBack();

    /// <summary>
    /// 取消组合形状
    /// 对应 Shape.Ungroup 方法
    /// </summary>
    /// <returns>取消组合后的形状集合</returns>
    IExcelShapeRange? Ungroup();

    /// <summary>
    /// 应用自动调整选项
    /// 对应 Shape.Apply 方法
    /// </summary>
    void Apply();

    /// <summary>
    /// 复制形状的格式
    /// 对应 Shape.PickUp 方法
    /// </summary>
    void PickUp();

    #endregion

    #region 层次结构

    /// <summary>
    /// 获取形状所在的区域对象（如果适用）
    /// 对应 Shape.TopLeftCell 属性
    /// </summary>
    IExcelRange? TopLeftCell { get; }

    /// <summary>
    /// 获取形状所在的区域对象（如果适用）
    /// 对应 Shape.BottomRightCell 属性
    /// </summary>
    IExcelRange? BottomRightCell { get; }

    /// <summary>
    /// 获取形状所在的图表对象（如果是图表）
    /// 对应 Shape.Chart 属性
    /// </summary>
    IExcelChart? Chart { get; }

    #endregion
}