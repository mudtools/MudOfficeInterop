//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示形状节点的集合，这些节点定义了自由形状的几何结构。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointShapeNodes : IOfficeObject<IPowerPointShapeNodes, MsPowerPoint.ShapeNodes>, IEnumerable<IPowerPointShapeNode?>, IDisposable
{
    /// <summary>
    /// 获取创建此形状节点集合的应用程序。
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
    /// 获取形状节点集合的父对象。
    /// </summary>
    /// <value>父对象，通常是自由形状。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中形状节点的数量。
    /// </summary>
    /// <value>集合中节点的总数。</value>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的形状节点。
    /// </summary>
    /// <param name="index">要获取的节点的索引。</param>
    /// <returns>指定索引处的形状节点。</returns>
    IPowerPointShapeNode? this[object index] { get; }

    /// <summary>
    /// 通过索引获取集合中的形状节点。
    /// </summary>
    /// <param name="index">要获取的节点的索引。</param>
    /// <returns>指定索引处的形状节点。</returns>
    IPowerPointShapeNode? this[string index] { get; }

    /// <summary>
    /// 删除指定索引处的形状节点。
    /// </summary>
    /// <param name="index">要删除的节点的索引（从1开始）。</param>
    void Delete(int index);

    /// <summary>
    /// 在指定索引处插入一个新的形状节点。
    /// </summary>
    /// <param name="index">要插入新节点的位置索引（从1开始）。</param>
    /// <param name="segmentType">新节点的线段类型。</param>
    /// <param name="editingType">新节点的编辑类型。</param>
    /// <param name="x1">第一个控制点的X坐标（磅）。</param>
    /// <param name="y1">第一个控制点的Y坐标（磅）。</param>
    /// <param name="x2">第二个控制点的X坐标（磅），对于某些线段类型可选。</param>
    /// <param name="y2">第二个控制点的Y坐标（磅），对于某些线段类型可选。</param>
    /// <param name="x3">第三个控制点的X坐标（磅），对于某些线段类型可选。</param>
    /// <param name="y3">第三个控制点的Y坐标（磅），对于某些线段类型可选。</param>
    void Insert(int index, [ComNamespace("MsCore")] MsoSegmentType segmentType, [ComNamespace("MsCore")] MsoEditingType editingType, float x1, float y1, float x2 = 0f, float y2 = 0f, float x3 = 0f, float y3 = 0f);

    /// <summary>
    /// 设置指定索引处节点的编辑类型。
    /// </summary>
    /// <param name="index">要设置的节点的索引（从1开始）。</param>
    /// <param name="editingType">要应用的编辑类型。</param>
    void SetEditingType(int index, [ComNamespace("MsCore")] MsoEditingType editingType);

    /// <summary>
    /// 设置指定索引处节点的位置。
    /// </summary>
    /// <param name="index">要设置的节点的索引（从1开始）。</param>
    /// <param name="x1">新的X坐标（磅）。</param>
    /// <param name="y1">新的Y坐标（磅）。</param>
    void SetPosition(int index, float x1, float y1);

    /// <summary>
    /// 设置指定索引处节点的线段类型。
    /// </summary>
    /// <param name="index">要设置的节点的索引（从1开始）。</param>
    /// <param name="segmentType">要应用的线段类型。</param>
    void SetSegmentType(int index, [ComNamespace("MsCore")] MsoSegmentType segmentType);
}