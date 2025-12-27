//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ShapeRange 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ShapeRange 的安全访问和操作
/// ShapeRange 代表一个或多个形状对象的集合
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelShapeRange : IOfficeObject<IExcelShapeRange>, IEnumerable<IExcelShape?>, IDisposable
{
    #region 基础属性

    /// <summary>
    /// 获取形状区域中的形状数量
    /// 对应 ShapeRange.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的形状对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">形状索引（从1开始）</param>
    /// <returns>形状对象</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelShape? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的形状对象
    /// </summary>
    /// <param name="name">形状名称</param>
    /// <returns>形状对象</returns>
    [ComPropertyWrap(NeedConvert = true)]
    IExcelShape? this[string name] { get; }

    /// <summary>
    /// 获取形状区域的名称
    /// 对应 ShapeRange.Name 属性
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状区域所在的父对象（通常是工作表）
    /// 对应 ShapeRange.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取形状区域的ID
    /// 对应 ShapeRange.ID 属性
    /// </summary>
    int ID { get; }

    #endregion

    #region 位置和大小属性

    /// <summary>
    /// 获取形状区域的左边距
    /// 对应 ShapeRange.Left 属性
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取形状区域的顶边距
    /// 对应 ShapeRange.Top 属性
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取形状区域的宽度
    /// 对应 ShapeRange.Width 属性
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取形状区域的高度
    /// 对应 ShapeRange.Height 属性
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取形状区域的旋转角度
    /// 对应 ShapeRange.Rotation 属性
    /// </summary>
    float Rotation { get; set; }

    #endregion

    #region 可见性和状态

    /// <summary>
    /// 获取或设置形状区域是否可见
    /// 对应 ShapeRange.Visible 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Visible { get; set; }

    #endregion

    #region 格式设置

    /// <summary>
    /// 获取形状区域的填充格式对象
    /// 对应 ShapeRange.Fill 属性
    /// </summary>
    IExcelFillFormat? Fill { get; }

    /// <summary>
    /// 获取形状区域的线条格式对象
    /// 对应 ShapeRange.Line 属性
    /// </summary>
    IExcelLineFormat? Line { get; }

    /// <summary>
    /// 获取形状区域的文本框架对象
    /// 对应 ShapeRange.TextFrame 属性
    /// </summary>
    IExcelTextFrame? TextFrame { get; }

    /// <summary>
    /// 获取形状区域的阴影格式对象
    /// 对应 ShapeRange.Shadow 属性
    /// </summary>
    IExcelShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取形状区域的三维格式对象
    /// 对应 ShapeRange.ThreeD 属性
    /// </summary>
    IExcelThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取形状区域的文本特效格式对象
    /// 对应 ShapeRange.TextEffect 属性
    /// </summary>
    IExcelTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取形状区域的调整设置对象
    /// 对应 ShapeRange.Adjustments 属性
    /// </summary>
    IExcelAdjustments? Adjustments { get; }

    /// <summary>
    /// 获取形状区域的标注线格式对象
    /// 对应 ShapeRange.Callout 属性
    /// </summary>
    IExcelCalloutFormat? Callout { get; }

    /// <summary>
    /// 获取形状区域的连接线格式对象
    /// 对应 ShapeRange.ConnectorFormat 属性
    /// </summary>
    IExcelConnectorFormat? ConnectorFormat { get; }

    /// <summary>
    /// 获取形状区域的图片格式对象
    /// 对应 ShapeRange.PictureFormat 属性
    /// </summary>
    IExcelPictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取形状区域的柔边格式对象
    /// 对应 ShapeRange.SoftEdge 属性
    /// </summary>
    IOfficeSoftEdgeFormat? SoftEdge { get; }

    /// <summary>
    /// 获取形状区域的发光格式对象
    /// 对应 ShapeRange.Glow 属性
    /// </summary>
    IOfficeGlowFormat? Glow { get; }

    /// <summary>
    /// 获取形状区域的倒影格式对象
    /// 对应 ShapeRange.Reflection 属性
    /// </summary>
    IOfficeReflectionFormat? Reflection { get; }

    /// <summary>
    /// 获取组合形状中的单个形状对象集合
    /// 对应 ShapeRange.GroupItems 属性
    /// </summary>
    IExcelGroupShapes? GroupItems { get; }

    /// <summary>
    /// 获取形状区域的节点对象集合
    /// 对应 ShapeRange.Nodes 属性
    /// </summary>
    IExcelShapeNodes? Nodes { get; }

    /// <summary>
    /// 获取形状区域的父组合形状对象
    /// 对应 ShapeRange.ParentGroup 属性
    /// </summary>
    IExcelShape? ParentGroup { get; }

    /// <summary>
    /// 获取或设置自动形状类型
    /// 对应 ShapeRange.AutoShapeType 属性
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoShapeType AutoShapeType { get; set; }

    /// <summary>
    /// 获取或设置形状区域的标题
    /// 对应 ShapeRange.Title 属性
    /// </summary>
    string Title { get; set; }

    /// <summary>
    /// 获取形状类型
    /// 对应 ShapeRange.Type 属性
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeType Type { get; }

    /// <summary>
    /// 获取或设置形状的黑白显示模式
    /// 对应 ShapeRange.BlackWhiteMode 属性
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBlackWhiteMode BlackWhiteMode { get; set; }

    /// <summary>
    /// 获取或设置形状样式
    /// 对应 ShapeRange.ShapeStyle 属性
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoShapeStyleIndex ShapeStyle { get; set; }

    /// <summary>
    /// 获取或设置形状背景样式
    /// 对应 ShapeRange.BackgroundStyle 属性
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取连接点数量
    /// 对应 ShapeRange.ConnectionSiteCount 属性
    /// </summary>
    int ConnectionSiteCount { get; }

    /// <summary>
    /// 获取一个值，该值指示形状是否为连接符
    /// 对应 ShapeRange.Connector 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Connector { get; }

    /// <summary>
    /// 获取一个值，该值指示形状是否包含图表
    /// 对应 ShapeRange.HasChart 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasChart { get; }

    /// <summary>
    /// 获取一个值，该值指示形状是否水平翻转
    /// 对应 ShapeRange.HorizontalFlip 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HorizontalFlip { get; }

    /// <summary>
    /// 获取一个值，该值指示形状是否垂直翻转
    /// 对应 ShapeRange.VerticalFlip 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool VerticalFlip { get; }

    /// <summary>
    /// 获取或设置是否锁定形状的纵横比
    /// 对应 ShapeRange.LockAspectRatio 属性
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool LockAspectRatio { get; set; }

    /// <summary>
    /// 获取形状在 Z 轴上的位置
    /// 对应 ShapeRange.ZOrderPosition 属性
    /// </summary>
    int ZOrderPosition { get; }
    #endregion


    #region 选择和操作

    /// <summary>
    /// 选择形状区域
    /// 对应 ShapeRange.Select 方法
    /// </summary>
    /// <param name="replace">true表示替换当前选择，false表示添加到当前选择</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除形状区域中的所有形状
    /// 对应 ShapeRange.Delete 方法
    /// </summary>
    void Delete();

    /// <summary>
    /// 应用自动调整选项
    /// 对应 ShapeRange.Apply 方法
    /// </summary>
    void Apply();

    /// <summary>
    /// 复制形状区域的格式
    /// 对应 ShapeRange.PickUp 方法
    /// </summary>
    void PickUp();

    #endregion


    #region 排列和布局

    /// <summary>
    /// 按指定比例缩放形状区域的高度
    /// </summary>
    /// <param name="factor">缩放因子，1.0表示原始大小，小于1.0表示缩小，大于1.0表示放大</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小进行缩放。true表示相对于原始大小，false表示相对于当前大小</param>
    /// <param name="scale">缩放参考点，指定从哪个位置开始缩放，默认为null</param>
    void ScaleHeight(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom? scale = null);

    /// <summary>
    /// 按指定比例缩放形状区域的宽度
    /// </summary>
    /// <param name="factor">缩放因子，1.0表示原始大小，小于1.0表示缩小，大于1.0表示放大</param>
    /// <param name="relativeToOriginalSize">是否相对于原始大小进行缩放。true表示相对于原始大小，false表示相对于当前大小</param>
    /// <param name="scale">缩放参考点，指定从哪个位置开始缩放，默认为null</param>
    void ScaleWidth(float factor, [ConvertTriState] bool relativeToOriginalSize, [ComNamespace("MsCore")] MsoScaleFrom? scale = null);


    /// <summary>
    /// 对齐形状区域中的形状
    /// </summary>
    /// <param name="alignment">对齐方式</param>
    /// <param name="relativeTo">相对对象</param>
    void Align([ComNamespace("MsCore")] MsoAlignCmd alignment, [ConvertTriState] bool relativeTo = false);

    /// <summary>
    /// 将形状区域中所有形状设置为默认属性
    /// </summary>
    void SetShapesDefaultProperties();

    /// <summary>
    /// 翻转形状区域中的形状
    /// </summary>
    /// <param name="FlipCmd">翻转方向，水平或垂直</param>
    void Flip([ComNamespace("MsCore")] MsoFlipCmd FlipCmd);

    /// <summary>
    /// 复制形状区域中的所有形状
    /// </summary>
    /// <returns>复制后的形状区域对象</returns>
    IExcelShapeRange? Duplicate();

    /// <summary>
    /// 增量调整形状区域的水平位置
    /// </summary>
    /// <param name="Increment">水平位置增量</param>
    void IncrementLeft(float Increment);

    /// <summary>
    /// 增量调整形状区域的垂直位置
    /// </summary>
    /// <param name="Increment">垂直位置增量</param>
    void IncrementTop(float Increment);

    /// <summary>
    /// 增量调整形状区域的旋转角度
    /// </summary>
    /// <param name="Increment">旋转角度增量</param>
    void IncrementRotation(float Increment);

    /// <summary>
    /// 重新路由连接线
    /// </summary>
    void RerouteConnections();

    /// <summary>
    /// 重新组合形状区域中的形状
    /// </summary>
    /// <returns>重新组合后的形状对象</returns>
    IExcelShape? Regroup();

    /// <summary>
    /// 设置形状区域中形状的堆叠顺序（Z轴顺序）
    /// </summary>
    /// <param name="ZOrderCmd">Z轴顺序命令</param>
    void ZOrder([ComNamespace("MsCore")] MsoZOrderCmd ZOrderCmd);

    /// <summary>
    /// 从左侧裁剪画布
    /// </summary>
    /// <param name="Increment">裁剪增量</param>
    void CanvasCropLeft(float Increment);

    /// <summary>
    /// 从顶部裁剪画布
    /// </summary>
    /// <param name="Increment">裁剪增量</param>
    void CanvasCropTop(float Increment);

    /// <summary>
    /// 从右侧裁剪画布
    /// </summary>
    /// <param name="Increment">裁剪增量</param>
    void CanvasCropRight(float Increment);

    /// <summary>
    /// 从底部裁剪画布
    /// </summary>
    /// <param name="Increment">裁剪增量</param>
    void CanvasCropBottom(float Increment);

    /// <summary>
    /// 分布形状区域中的形状
    /// </summary>
    /// <param name="distribution">分布方式</param>
    /// <param name="relativeTo">是否相对于当前位置进行分布</param> 
    void Distribute([ComNamespace("MsCore")] MsoDistributeCmd distribution, [ConvertTriState] bool relativeTo = false);

    #endregion

    #region 组合操作

    /// <summary>
    /// 组合形状区域中的所有形状
    /// 对应 ShapeRange.Group 方法
    /// </summary>
    /// <returns>组合后的形状对象</returns>
    IExcelShape? Group();

    /// <summary>
    /// 取消组合形状区域中的组合形状
    /// 对应 ShapeRange.Ungroup 方法
    /// </summary>
    /// <returns>取消组合后的形状区域</returns>
    IExcelShapeRange? Ungroup();

    #endregion

    #region 层次结构
    /// <summary>
    /// 获取形状区域所在的图表对象（如果是图表）
    /// 对应 ShapeRange.Chart 属性
    /// </summary>
    IExcelChart? Chart { get; }
    #endregion   
}