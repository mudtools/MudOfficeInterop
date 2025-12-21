//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 图表数据标签集合的封装接口。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordDataLabels : IEnumerable<IWordDataLabel?>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据标签数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取数据标签。
    /// </summary>
    IWordDataLabel? this[int index] { get; }

    /// <summary>
    /// 获取或设置是否自动文本。
    /// </summary>
    bool AutoText { get; set; }

    /// <summary>
    /// 获取或设置是否显示图例项标示。
    /// </summary>
    bool ShowLegendKey { get; set; }

    /// <summary>
    /// 获取或设置是否显示值。
    /// </summary>
    bool ShowValue { get; set; }

    /// <summary>
    /// 获取或设置是否显示分类名称。
    /// </summary>
    bool ShowCategoryName { get; set; }

    /// <summary>
    /// 获取或设置是否显示系列名称。
    /// </summary>
    bool ShowSeriesName { get; set; }

    /// <summary>
    /// 获取或设置是否显示百分比。
    /// </summary>
    bool ShowPercentage { get; set; }

    /// <summary>
    /// 获取或设置是否显示气泡大小。
    /// </summary>
    bool ShowBubbleSize { get; set; }

    /// <summary>
    /// 获取或设置数据标签位置。
    /// </summary>
    XlDataLabelPosition Position { get; set; }

    /// <summary>
    /// 获取或设置水平对齐方式。
    /// </summary>
    XlConstants HorizontalAlignment { get; set; }

    /// <summary>
    /// 获取或设置垂直对齐方式。
    /// </summary>
    XlConstants VerticalAlignment { get; set; }

    /// <summary>
    /// 获取或设置是否自动缩放字体。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool AutoScaleFont { get; set; }

    /// <summary>
    /// 获取或设置是否显示数据标签范围。
    /// </summary>
    bool ShowRange { get; set; }

    /// <summary>
    /// 获取或设置数据标签分隔符。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string Separator { get; set; }

    /// <summary>
    /// 获取或设置数据标签类型。
    /// </summary>
    XlDataLabelsType Type { get; set; }

    /// <summary>
    /// 获取或设置本地化数字格式。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    string NumberFormatLocal { get; set; }

    /// <summary>
    /// 获取或设置数字格式是否链接到单元格。
    /// </summary>
    bool NumberFormatLinked { get; set; }

    /// <summary>
    /// 获取或设置数字格式。
    /// </summary>
    string NumberFormat { get; set; }

    /// <summary>
    /// 获取或设置读取顺序。
    /// </summary>
    int ReadingOrder { get; set; }

    /// <summary>
    /// 获取或设置数据标签方向。
    /// </summary>
    XlOrientation Orientation { get; set; }

    /// <summary>
    /// 获取字体格式。
    /// </summary>
    IWordChartFont? Font { get; }

    /// <summary>
    /// 获取内部区域格式。
    /// </summary>
    IWordInterior? Interior { get; }

    /// <summary>
    /// 获取填充格式。
    /// </summary>
    IWordChartFillFormat? Fill { get; }

    /// <summary>
    /// 获取边框格式。
    /// </summary>
    IWordChartBorder? Border { get; }

    /// <summary>
    /// 获取格式对象。
    /// </summary>
    IWordChartFormat? Format { get; }

    /// <summary>
    /// 删除所有数据标签。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择数据标签集合。
    /// </summary>
    void Select();

    /// <summary>
    /// 将指定数据标签的内容和格式传播到系列中的所有其他数据标签。
    /// </summary>
    /// <param name="Index">要传播的数据标签的 DataLabels 集合中的索引号。</param>
    void Propagate(int Index);
}