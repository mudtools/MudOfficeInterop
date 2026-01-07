
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示自由曲线在构建过程中的几何形状。
/// 此接口提供在创建自由曲线时动态添加节点和线段的功能，最终可将构建的几何形状转换为实际的图形对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFreeformBuilder : IOfficeObject<IWordFreeformBuilder, MsWord.FreeformBuilder>, IDisposable
{
    /// <summary>
    /// 获取此标题样式所属的 Word 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此标题样式的父对象（通常是 <see cref="IWordHeadingStyles"/> 集合）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 在正在创建的自由曲线末尾插入新线段，并添加定义该线段的节点。
    /// </summary>
    /// <param name="segmentType">要添加的线段类型。</param>
    /// <param name="editingType">顶点的编辑属性。如果 segmentType 为 msoSegmentLine，则 editingType 必须为 msoEditingAuto。</param>
    /// <param name="x1">
    /// 如果新线段的编辑类型为 msoEditingAuto，则此参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。
    /// 如果新节点的编辑类型为 msoEditingCorner，则此参数指定从文档左上角到新线段第一个控制点的水平距离（以磅为单位）。
    /// </param>
    /// <param name="y1">
    /// 如果新线段的编辑类型为 msoEditingAuto，则此参数指定从文档左上角到新线段终点的垂直距离（以磅为单位）。
    /// 如果新节点的编辑类型为 msoEditingCorner，则此参数指定从文档左上角到新线段第一个控制点的垂直距离（以磅为单位）。
    /// </param>
    /// <param name="x2">
    /// 如果新线段的编辑类型为 msoEditingCorner，则此参数指定从文档左上角到新线段第二个控制点的水平距离（以磅为单位）。
    /// 如果新线段的编辑类型为 msoEditingAuto，则无需为此参数指定值。
    /// </param>
    /// <param name="y2">
    /// 如果新线段的编辑类型为 msoEditingCorner，则此参数指定从文档左上角到新线段第二个控制点的垂直距离（以磅为单位）。
    /// 如果新线段的编辑类型为 msoEditingAuto，则无需为此参数指定值。
    /// </param>
    /// <param name="x3">
    /// 如果新线段的编辑类型为 msoEditingCorner，则此参数指定从文档左上角到新线段终点的水平距离（以磅为单位）。
    /// 如果新线段的编辑类型为 msoEditingAuto，则无需为此参数指定值。
    /// </param>
    /// <param name="y3">
    /// 如果新线段的编辑类型为 msoEditingCorner，则此参数指定从文档左上角到新线段终点的垂直距离（以磅为单位）。
    /// 如果新线段的编辑类型为 msoEditingAuto，则无需为此参数指定值。
    /// </param>
    void AddNodes([ComNamespace("MsCore")] MsoSegmentType segmentType, [ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1, float x2 = 0f, float y2 = 0f, float x3 = 0f, float y3 = 0f);

    /// <summary>
    /// 创建一个具有指定对象几何特征的形状。返回表示新形状的 Shape 对象。
    /// </summary>
    /// <param name="anchor">表示形状绑定的文本范围的 Range 对象。
    /// 如果指定了 anchor，则定位点位于定位范围第一段的开头。
    /// 如果省略此参数，则自动选择定位范围，形状相对于页面的顶部和左边缘定位。</param>
    /// <returns>新创建的形状对象。</returns>
    IWordShape? ConvertToShape(IWordRange? anchor = null);
}