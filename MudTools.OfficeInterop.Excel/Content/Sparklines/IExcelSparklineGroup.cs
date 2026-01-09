//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示一组迷你图。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelSparklineGroup : IEnumerable<IExcelSparkline?>, IOfficeObject<IExcelSparklineGroup, MsExcel.SparklineGroup>, IDisposable
{

    /// <summary>
    /// 获取所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取所属的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取迷你图组中的迷你图数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取或设置表示迷你图组位置的 Range 对象。
    /// </summary>
    IExcelRange? Location { get; set; }

    /// <summary>
    /// 获取或设置包含迷你图组源数据的区域。
    /// </summary>
    string SourceData { get; set; }

    /// <summary>
    /// 获取或设置迷你图组的日期范围。
    /// </summary>
    string DateRange { get; set; }

    /// <summary>
    /// 修改迷你图组的位置。
    /// </summary>
    /// <param name="location">表示迷你图组新位置的 Range 对象。</param>
    void ModifyLocation(IExcelRange location);

    /// <summary>
    /// 修改迷你图组的源数据。
    /// </summary>
    /// <param name="sourceData">表示新源数据的区域。</param>
    void ModifySourceData(string sourceData);

    /// <summary>
    /// 同时修改迷你图组的位置和源数据。
    /// </summary>
    /// <param name="location">表示迷你图组新位置的 Range 对象。</param>
    /// <param name="sourceData">表示新源数据的区域。</param>
    void Modify(IExcelRange location, string sourceData);

    /// <summary>
    /// 修改迷你图组的日期范围。
    /// </summary>
    /// <param name="dateRange">新的日期范围。</param>
    void ModifyDateRange(string dateRange);

    /// <summary>
    /// 删除迷你图组。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取或设置迷你图组的类型。
    /// </summary>
    XlSparkType Type { get; set; }

    /// <summary>
    /// 获取表示迷你图组主系列颜色的 FormatColor 对象。
    /// </summary>
    IExcelFormatColor? SeriesColor { get; }

    /// <summary>
    /// 获取迷你图组中的点属性。
    /// </summary>
    IExcelSparkPoints? Points { get; }

    /// <summary>
    /// 获取关联的 SparkAxes 对象。
    /// </summary>
    IExcelSparkAxes? Axes { get; }

    /// <summary>
    /// 获取或设置图表中空白单元格的绘制方式。
    /// </summary>
    XlDisplayBlanksAs DisplayBlanksAs { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在迷你图组中绘制隐藏单元格。
    /// </summary>
    bool DisplayHidden { get; set; }

    /// <summary>
    /// 获取或设置迷你图组中迷你线的粗细。
    /// </summary>
    object LineWeight { get; set; }

    /// <summary>
    /// 获取或设置当基于正方形区域的数据时如何绘制迷你图。
    /// </summary>
    XlSparklineRowCol PlotBy { get; set; }
}