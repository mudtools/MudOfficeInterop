

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示图表系列中所有数据标签的集合。
/// 此接口是对 Microsoft.Office.Interop.Excel.DataLabels COM 对象的封装。
/// </summary>
public interface IExcelDataLabels : IEnumerable<IExcelDataLabel>, IDisposable
{
    /// <summary>
    /// 获取集合中的数据标签总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取集合中指定索引处的数据标签。
    /// 索引从 1 开始。
    /// </summary>
    /// <param name="index">数据标签的索引（从1开始）。</param>
    /// <returns>指定索引处的 <see cref="IExcelDataLabel"/> 对象。</returns>
    IExcelDataLabel? this[int index] { get; }

    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在所有数据标签中显示图例项标示。
    /// </summary>
    bool ShowLegendKey { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在所有数据标签中显示百分比值。
    /// </summary>
    bool ShowPercentage { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在所有数据标签中显示系列名称。
    /// </summary>
    bool ShowSeriesName { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在所有数据标签中显示分类名称。
    /// </summary>
    bool ShowCategoryName { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在所有数据标签中显示数值。
    /// </summary>
    bool ShowValue { get; set; }

    /// <summary>
    /// 获取或设置所有数据标签的位置。
    /// </summary>
    XlDataLabelPosition Position { get; set; }

    /// <summary>
    /// 获取或设置所有数据标签的数字格式代码。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示所有数据标签的数字格式是否为链接源格式。
    /// </summary>
    bool NumberFormatLinked { get; set; }

    /// <summary>
    /// 获取一个 <see cref="IExcelChartFormat"/> 对象，用于设置所有数据标签的整体格式。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelBorder"/> 对象，用于设置所有数据标签的边框。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelInterior"/> 对象，用于设置所有数据标签的内部填充。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelFont"/> 对象，用于设置所有数据标签的字体。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置所有数据标签的水平对齐方式。
    /// </summary>
    object? HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置所有数据标签的垂直对齐方式。
    /// </summary>
    object? VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置所有数据标签文本的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置所有数据标签文本的方向（角度）。
    /// </summary>
    object? Orientation { get; set; }

    /// <summary>
    /// 删除集合中的所有数据标签。
    /// </summary>
    void Delete();

    /// <summary>
    /// 将单个数据标签的内容和格式应用到系列中的所有其他数据标签。
    /// </summary>
    /// <param name="index">要传播的单个数据标签的索引。</param>
    void Propagate(int index);
}