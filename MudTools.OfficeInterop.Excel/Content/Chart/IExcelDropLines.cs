
namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// 表示 Excel 图表中从数据点垂直到分类轴的“垂线”的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.DropLines
/// 常用于折线图、面积图等，增强数据可读性。
/// </summary>
public interface IExcelDropLines : IDisposable
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
    /// 获取垂线的边框格式（控制线条颜色、粗细、样式）。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取垂线的图表格式对象，用于控制垂线的填充、线条、阴影等高级格式设置。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 选中此垂线（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此垂线（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}