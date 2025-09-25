
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 图表中坐标轴标题的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.AxisTitle
/// 用于设置标题文本、字体、对齐方式、方向、可见性等。
/// </summary>
public interface IExcelAxisTitle : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（通常是 Axis）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取坐标轴标题的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置坐标轴标题的显示文本。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题相对于图表左边的距离（单位：磅）。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题相对于图表顶部的距离（单位：磅）。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取坐标轴标题的宽度（单位：磅）。
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取坐标轴标题的高度（单位：磅）。
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取或设置坐标轴标题是否显示阴影效果。
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题是否包含在图表布局中。
    /// </summary>
    bool IncludeInLayout { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题在图表中的位置类型。
    /// </summary>
    XlChartElementPosition Position { get; set; }

    /// <summary>
    /// 获取坐标轴标题的内部填充格式。
    /// </summary>
    IExcelInterior? Interior { get; }

    /// <summary>
    /// 获取或设置坐标轴标题的公式（A1样式引用）。
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的公式（R1C1样式引用）。
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的本地化公式（A1样式引用）。
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的本地化公式（R1C1样式引用）。
    /// </summary>
    string FormulaR1C1Local { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取坐标轴标题的文本方向（角度或预设方向）。
    /// 使用 <see cref="XlOrientation"/> 枚举。
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的水平对齐方式。
    /// 使用 <see cref="XlHAlign"/> 枚举。
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置坐标轴标题的垂直对齐方式。
    /// 使用 <see cref="XlVAlign"/> 枚举。
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取坐标轴标题的字体格式。
    /// 返回封装后的 <see cref="IExcelFont"/> 接口。
    /// </summary>
    IExcelFont? Font { get; }

    /// <summary>
    /// 获取坐标轴标题的字符格式（用于高级文本格式）。
    /// 返回封装后的 <see cref="IExcelCharacters"/> 接口。
    /// </summary>
    IExcelCharacters? Characters { get; }

    /// <summary>
    /// 获取坐标轴标题的填充格式。
    /// 返回封装后的 <see cref="IExcelChartFillFormat"/> 接口。
    /// </summary>
    IExcelChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取坐标轴标题的边框格式。
    /// 返回封装后的 <see cref="IExcelBorder"/> 接口。
    /// </summary>
    IExcelBorder? Border { get; }

    /// <summary>
    /// 选中此坐标轴标题（激活并高亮显示）。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除此坐标轴标题（将其设为不可见，并从图表中移除）。
    /// </summary>
    void Delete();
}