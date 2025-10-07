
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 股票图中连接最高价和最低价的“高低线”的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.HiLoLines
/// 用于股票图（High-Low-Close 等）中显示价格波动范围。
/// </summary>
public interface IExcelHiLoLines : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Chart）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取高低线的边框格式（控制线条颜色、粗细、样式）。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取高低线的图表格式对象，用于控制线条的高级格式设置（如填充、线条样式、阴影等）。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 选中此高低线（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此高低线（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}