
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 股票图或柱形图中“下跌柱”的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.DownBars
/// 用于设置下跌柱（开盘价高于收盘价）的填充、边框、可见性等。
/// </summary>
public interface IExcelDownBars : IDisposable
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
    /// 获取下跌柱的边框格式。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取下跌柱的填充格式。
    /// 返回封装后的 <see cref="IExcelChartFillFormat"/> 接口。
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取下跌柱的内部填充格式。
    /// 返回封装后的 <see cref="IExcelInterior"/> 接口，用于设置下跌柱的背景色、图案等内部样式属性。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取下跌柱的图表格式对象。
    /// 返回封装后的 <see cref="IExcelChartFormat"/> 接口，用于设置下跌柱的通用格式属性，如填充、线条、阴影等。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 选中此下跌柱（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此下跌柱（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}