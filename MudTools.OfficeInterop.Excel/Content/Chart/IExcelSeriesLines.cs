
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 组合图表中系列连线（Series Lines）的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.SeriesLines
/// 用于控制连接主次坐标轴数据系列的连线样式（如柱形图与折线图之间的连线）。
/// </summary>
public interface IExcelSeriesLines : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Chart 或 SeriesCollection）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取系列连线的边框格式（用于设置颜色、线型、粗细等）。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取系列连线的格式设置对象，用于控制其外观样式（如填充、线条、阴影等）。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 选中此系列连线（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此系列连线（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}