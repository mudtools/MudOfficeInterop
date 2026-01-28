//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中形状范围的接口封装，用于同时操作多个形状。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsCore")]
[ItemIndex]
public interface IOfficeShapeRange : IOfficeObject<IOfficeShapeRange, MsCore.ShapeRange>, IEnumerable<IOfficeShape?>, IDisposable
{
    /// <summary>
    /// 获取形状范围中形状的数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取形状（从 1 开始）。
    /// </summary>
    /// <param name="index">形状索引。</param>
    /// <returns>对应的形状对象。</returns>
    IOfficeShape? this[int index] { get; }

    /// <summary>
    /// 通过名称获取形状。
    /// </summary>
    /// <param name="name">形状名称。</param>
    /// <returns>对应的形状对象。</returns>
    IOfficeShape? this[string name] { get; }

    /// <summary>
    /// 获取形状范围的名称。
    /// </summary>
    string Name { get; set; }

    /// <summary>
    /// 获取形状范围的高度（单位：磅）。
    /// </summary>
    float Height { get; set; }

    /// <summary>
    /// 获取形状范围的宽度（单位：磅）。
    /// </summary>
    float Width { get; set; }

    /// <summary>
    /// 获取形状范围的左边缘位置。
    /// </summary>
    float Left { get; set; }

    /// <summary>
    /// 获取形状范围的上边缘位置。
    /// </summary>
    float Top { get; set; }

    /// <summary>
    /// 获取形状范围的旋转角度。
    /// </summary>
    float Rotation { get; set; }

    /// <summary>
    /// 获取形状范围的填充格式。
    /// </summary>
    IOfficeFillFormat? Fill { get; }

    /// <summary>
    /// 获取形状范围的线条格式。
    /// </summary>
    IOfficeLineFormat? Line { get; }

    /// <summary>
    /// 获取形状范围的阴影格式。
    /// </summary>
    IOfficeShadowFormat? Shadow { get; }

    /// <summary>
    /// 获取一个值，该值指示形状是否为连接符。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool Connector { get; }

    /// <summary>
    /// 获取形状的连接符格式。
    /// </summary>
    IOfficeConnectorFormat? ConnectorFormat { get; }

    /// <summary>
    /// 获取形状的三维格式。
    /// </summary>
    IOfficeThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取形状的文字效果格式。
    /// </summary>
    IOfficeTextEffectFormat? TextEffect { get; }

    /// <summary>
    /// 获取形状的图片格式。
    /// </summary>
    IOfficePictureFormat? PictureFormat { get; }

    /// <summary>
    /// 获取形状范围的文本框架。
    /// </summary>
    IOfficeTextFrame? TextFrame { get; }

    /// <summary>
    /// 对形状范围进行分组。
    /// </summary>
    /// <returns>分组后的形状对象。</returns>
    IOfficeShape? Group();

    /// <summary>
    /// 取消形状范围的分组。
    /// </summary>
    /// <returns>取消分组后的形状范围。</returns>
    IOfficeShapeRange? Ungroup();

    /// <summary>
    /// 选择形状范围。
    /// </summary>
    /// <param name="replace">是否替换当前选择。</param>
    void Select(bool replace = true);

    /// <summary>
    /// 删除形状范围中的所有形状。
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制形状范围。
    /// </summary>
    void Copy();

    /// <summary>
    /// 剪切形状范围。
    /// </summary>
    void Cut();

    /// <summary>
    /// 水平翻转形状范围。
    /// </summary>
    /// <param name="flipCmd">翻转命令。</param>
    void Flip(MsoFlipCmd flipCmd);

    /// <summary>
    /// 设置形状范围的水平对齐方式。
    /// </summary>
    /// <param name="alignCmd">对齐命令。</param>
    /// <param name="relativeTo">相对于什么对齐。</param>
    void Align(MsoAlignCmd alignCmd, [ConvertTriState] bool relativeTo = false);

    /// <summary>
    /// 分布形状范围中的形状。
    /// </summary>
    /// <param name="distributeCmd">分布命令。</param>
    /// <param name="relativeTo">相对于什么分布。</param>
    void Distribute(MsoDistributeCmd distributeCmd, [ConvertTriState] bool relativeTo = false);

    /// <summary>
    /// 设置形状范围的Z轴顺序。
    /// </summary>
    /// <param name="zOrderCmd">Z轴顺序命令。</param>
    void ZOrder(MsoZOrderCmd zOrderCmd);

    /// <summary>
    /// 应用锁定比例缩放。
    /// </summary>
    /// <param name="scale">缩放比例。</param>
    /// <param name="scaleWidth">是否缩放宽度。</param>
    /// <param name="scaleHeight">是否缩放高度。</param>
    void ScaleHeight(float scale, [ConvertTriState] bool scaleWidth, MsoScaleFrom scaleHeight = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 应用锁定比例缩放。
    /// </summary>
    /// <param name="scale">缩放比例。</param>
    /// <param name="scaleWidth">是否缩放宽度。</param>
    /// <param name="scaleHeight">是否缩放高度。</param>
    void ScaleWidth(float scale, [ConvertTriState] bool scaleWidth, MsoScaleFrom scaleHeight = MsoScaleFrom.msoScaleFromTopLeft);

    /// <summary>
    /// 移动形状范围。
    /// </summary>
    /// <param name="deltaX">水平移动距离。</param>
    void IncrementLeft(float deltaX);

    /// <summary>
    /// 垂直移动形状范围。
    /// </summary>
    /// <param name="deltaY">垂直移动距离。</param>
    void IncrementTop(float deltaY);

    /// <summary>
    /// 旋转形状范围。
    /// </summary>
    /// <param name="increment">旋转角度增量。</param>
    void IncrementRotation(float increment);
}