//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示Excel图表中坐标轴的显示单位标签。
/// 当图表包含较大数值时，可以使用显示单位标签来简化坐标轴刻度值的显示，使图表更易读。
/// 例如，对于百万级别的数据，可以将坐标轴显示为1-50，并使用显示单位标签标明单位为"百万"。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelDisplayUnitLabel : IDisposable
{
    /// <summary>
    /// 获取显示单位标签的父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取与此显示单位标签关联的Excel应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取显示单位标签的边框格式。
    /// </summary>
    IExcelBorder Border { get; }

    /// <summary>
    /// 获取显示单位标签的内部填充格式。
    /// </summary>
    IExcelInterior Interior { get; }

    /// <summary>
    /// 获取显示单位标签的填充格式。
    /// </summary>
    IExcelChartFillFormat Fill { get; }

    /// <summary>
    /// 获取显示单位标签的字符格式。
    /// </summary>
    IExcelCharacters Characters { get; }

    /// <summary>
    /// 获取显示单位标签的字体格式。
    /// </summary>
    IExcelFont Font { get; }

    /// <summary>
    /// 获取显示单位标签的整体格式设置对象。
    /// </summary>
    IExcelChartFormat Format { get; }

    /// <summary>
    /// 获取或设置显示单位标签的水平对齐方式。
    /// </summary>
    XlHAlign HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的文字方向。
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的垂直对齐方式。
    /// </summary>
    XlVAlign VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置是否自动调整显示单位标签的字体大小。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签在图表中的位置。
    /// </summary>
    XlChartElementPosition Position { get; set; }

    /// <summary>
    /// 获取显示单位标签的名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取或设置显示单位标签的标题文本。
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签相对于图表左边缘的位置（单位：磅）。
    /// </summary>
    double Left { get; set; }

    /// <summary>
    /// 获取显示单位标签的高度（单位：磅）。
    /// </summary>
    double Height { get; }

    /// <summary>
    /// 获取或设置显示单位标签相对于图表上边缘的位置（单位：磅）。
    /// </summary>
    double Top { get; set; }

    /// <summary>
    /// 获取显示单位标签的宽度（单位：磅）。
    /// </summary>
    double Width { get; }

    /// <summary>
    /// 获取或设置是否显示显示单位标签的阴影效果。
    /// </summary>
    bool Shadow { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的公式（A1引用样式）。
    /// </summary>
    string Formula { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的公式（R1C1引用样式）。
    /// </summary>
    string FormulaR1C1 { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的本地化公式（A1引用样式）。
    /// </summary>
    string FormulaLocal { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的本地化公式（R1C1引用样式）。
    /// </summary>
    string FormulaR1C1Local { get; set; }

    /// <summary>
    /// 获取或设置显示单位标签的阅读顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 选中显示单位标签。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object Select();

    /// <summary>
    /// 删除显示单位标签。
    /// </summary>
    /// <returns>操作结果对象。</returns>
    object Delete();
}