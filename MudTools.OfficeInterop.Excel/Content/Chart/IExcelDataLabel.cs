
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示图表中单个数据点或趋势线的数据标签。
/// 此接口是对 Microsoft.Office.Interop.Excel.DataLabel COM 对象的封装。
/// </summary>
public interface IExcelDataLabel : IDisposable
{
    /// <summary>
    /// 获取该对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelApplication"/> 对象，该对象代表 Microsoft Excel 应用程序。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置数据标签的文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示数据标签是否自动生成适当的文本。
    /// </summary>
    bool AutoText { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在数据标签中显示图例项标示。
    /// </summary>
    bool ShowLegendKey { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在数据标签中显示百分比值。
    /// </summary>
    bool ShowPercentage { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在数据标签中显示系列名称。
    /// </summary>
    bool ShowSeriesName { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在数据标签中显示分类名称。
    /// </summary>
    bool ShowCategoryName { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在数据标签中显示数值。
    /// </summary>
    bool ShowValue { get; set; }

    /// <summary>
    /// 获取或设置数据标签的位置。
    /// </summary>
    XlDataLabelPosition Position { get; set; }

    /// <summary>
    /// 获取或设置数据标签的数字格式代码。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示数字格式是否为链接源格式。
    /// </summary>
    bool NumberFormatLinked { get; set; }

    /// <summary>
    /// 获取一个 <see cref="IExcelChartFormat"/> 对象，用于设置数据标签的整体格式。
    /// </summary>
    IExcelChartFormat? Format { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelBorder"/> 对象，用于设置数据标签的边框。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelInterior"/> 对象，用于设置数据标签的内部填充。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取一个 <see cref="IExcelFont"/> 对象，用于设置数据标签的字体。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取或设置数据标签的水平对齐方式。
    /// </summary>
    object? HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置数据标签的垂直对齐方式。
    /// </summary>
    object? VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置数据标签文本的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置数据标签文本的方向（角度）。
    /// </summary>
    object? Orientation { get; set; }

    /// <summary>
    /// 获取或设置数据标签的左边缘位置（以磅为单位）。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置数据标签的顶部位置（以磅为单位）。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取或设置数据标签的宽度（以磅为单位）。
    /// </summary>
    double Width { get; set; }

    /// <summary>
    /// 获取或设置数据标签的高度（以磅为单位）。
    /// </summary>
    double Height { get; set; }

    /// <summary>
    /// 删除该数据标签。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择该数据标签。
    /// </summary>
    void Select();
}
