namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示一个Excel自由多边形形状构建器接口，用于创建自定义的自由形状
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelFreeformBuilder : IOfficeObject<IExcelFreeformBuilder, MsExcel.FreeformBuilder>, IDisposable
{
    /// <summary>
    /// 获取对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的Excel应用程序对象
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 向自由多边形形状中添加节点
    /// </summary>
    /// <param name="segmentType">线段类型，指定要添加的线段类型</param>
    /// <param name="editingType">编辑类型，指定节点的编辑类型（角点或平滑点等）</param>
    /// <param name="x1">第一个节点的X坐标</param>
    /// <param name="y1">第一个节点的Y坐标</param>
    /// <param name="x2">第二个节点的X坐标（可选）</param>
    /// <param name="y2">第二个节点的Y坐标（可选）</param>
    /// <param name="x3">第三个节点的X坐标（可选）</param>
    /// <param name="y3">第三个节点的Y坐标（可选）</param>
    void AddNodes([ComNamespace("MsCore")] MsoSegmentType segmentType, [ComNamespace("MsCore")] MsoEditingType editingType,
                    float x1, float y1, float? x2 = null, float? y2 = null, float? x3 = null, float? y3 = null);

    /// <summary>
    /// 将自由多边形形状构建器转换为实际的形状对象
    /// </summary>
    /// <returns>转换后的Excel形状对象</returns>
    IExcelShape? ConvertToShape();
}