//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示指定自由曲线中所有 ShapeNode 对象的集合。
/// 此集合提供对自由曲线节点几何形状和编辑属性的访问，以及添加、删除和修改节点的功能。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordShapeNodes : IEnumerable<IWordShapeNode?>, IOfficeObject<IWordShapeNodes, MsWord.ShapeNodes>, IDisposable
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
    /// 获取集合中的节点数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取指定的形状节点。
    /// </summary>
    /// <param name="index">节点的序号位置或表示节点名称的字符串。</param>
    /// <returns>指定索引处的形状节点对象。</returns>
    IWordShapeNode? this[int index] { get; }

    /// <summary>
    /// 通过索引获取指定的形状节点。
    /// </summary>
    /// <param name="name">节点的序号位置或表示节点名称的字符串。</param>
    /// <returns>指定索引处的形状节点对象。</returns>
    IWordShapeNode? this[string name] { get; }

    /// <summary>
    /// 删除指定索引处的形状节点。
    /// </summary>
    /// <param name="index">要删除的节点序号。</param>
    void Delete(int index);

    /// <summary>
    /// 设置指定节点的编辑类型。
    /// </summary>
    /// <param name="index">要设置编辑类型的节点序号。</param>
    /// <param name="editingType">顶点的编辑属性。</param>
    void SetEditingType(int index, [ComNamespace("MsCore")] MsoEditingType editingType);

    /// <summary>
    /// 设置指定节点的位置。
    /// </summary>
    /// <param name="index">要设置位置的节点序号。</param>
    /// <param name="x1">新节点相对于文档左上角的水平位置（以磅为单位）。</param>
    /// <param name="y1">新节点相对于文档左上角的垂直位置（以磅为单位）。</param>
    void SetPosition(int index, float x1, float y1);

    /// <summary>
    /// 设置指定节点之后线段的类型。
    /// </summary>
    /// <param name="index">要设置线段类型的节点序号。</param>
    /// <param name="segmentType">指定线段是直线还是曲线。</param>
    void SetSegmentType(int index, [ComNamespace("MsCore")] MsoSegmentType segmentType);

    /// <summary>
    /// 在自由曲线中插入一个节点。
    /// </summary>
    /// <param name="index">在此节点序号之后插入新节点。</param>
    /// <param name="segmentType">连接插入节点与相邻节点的线条类型。</param>
    /// <param name="editingType">插入节点的编辑属性。</param>
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
    void Insert(int index, [ComNamespace("MsCore")] MsoSegmentType segmentType, [ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1, float x2 = 0f, float y2 = 0f, float x3 = 0f, float y3 = 0f);
}