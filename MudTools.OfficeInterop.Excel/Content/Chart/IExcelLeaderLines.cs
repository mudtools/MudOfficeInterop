namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 图表中的引导线对象
/// 引导线是连接数据标签与其对应数据点的线条，在数据标签偏离数据点较远时非常有用
/// 此接口封装了对 Excel 图表引导线对象的操作
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelLeaderLines : IOfficeObject<IExcelLeaderLines>, IDisposable
{
    /// <summary>
    /// 获取引导线对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取包含该引导线对象的 Excel 应用程序对象
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取引导线的边框格式对象，可用于设置引导线的样式、颜色和粗细等属性
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取引导线的图表格式对象，可用于设置引导线的整体格式属性
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 删除图表中的引导线
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择引导线对象
    /// </summary>
    void Select();
}