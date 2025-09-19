
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示自由形状中所有路径节点的集合，支持遍历、索引访问和节点操作。
/// </summary>
public interface IExcelShapeNodes : IEnumerable<IExcelShapeNode>, IDisposable
{
    /// <summary>
    /// 获取集合中节点的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的节点。
    /// </summary>
    /// <param name="index">节点索引（1-based）</param>
    /// <returns>对应的节点对象</returns>
    IExcelShapeNode this[int index] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Shape）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 在指定索引位置插入一个新节点。
    /// </summary>
    /// <param name="index">插入位置（从 1 开始）</param>
    /// <param name="segmentType">节点类型（直线或曲线）</param>
    /// <param name="editingType">节点类型（角点或曲线）</param>
    /// <param name="x1">节点 X 坐标</param>
    /// <param name="y1">节点 Y 坐标</param>
    /// <param name="x2">第一个控制点 X 偏移（仅曲线节点需要）</param>
    /// <param name="y2">第一个控制点 Y 偏移（仅曲线节点需要）</param>
    /// <param name="x3">第二个控制点 X 偏移（仅曲线节点需要）</param>
    /// <param name="y3">第二个控制点 Y 偏移（仅曲线节点需要）</param>
    /// <returns>新创建的节点对象</returns>
    void Insert(
       int index,
       MsoSegmentType segmentType,
       MsoEditingType editingType,
       float x1, float y1,
       float x2 = 0, float y2 = 0,
       float x3 = 0, float y3 = 0);

    /// <summary>
    /// 设置指定索引节点的属性。
    /// </summary>
    /// <param name="index">节点索引（1-based）</param>
    /// <param name="x1">节点 X 坐标</param>
    /// <param name="y1">节点 Y 坐标</param>
    void SetPosition(int index, float x1, float y1);
}