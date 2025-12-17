
namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 股票图或柱形图中“上涨柱”的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.UpBars
/// 用于设置上涨柱（开盘价低于收盘价）的填充、边框、可见性等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelUpBars : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Chart）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取上涨柱的边框格式。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取上涨柱对象的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取上涨柱的填充格式。
    /// 返回封装后的 <see cref="IExcelChartFillFormat"/> 接口。
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取上涨柱的内部区域格式（背景色、图案等）。
    /// 返回封装后的 <see cref="IExcelInterior"/> 接口。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取上涨柱的图表格式设置（包括大小、位置和效果等格式）。
    /// 返回封装后的 <see cref="IExcelChartFormat"/> 接口。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 选中此上涨柱（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此上涨柱（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}