//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作表中所有切片器的集合，支持遍历和索引访问。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSlicers : IOfficeObject<IExcelSlicers>, IEnumerable<IExcelSlicer?>, IDisposable
{
    /// <summary>
    /// 获取集合中切片器的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）取指定的切片器。
    /// </summary>
    /// <param name="index">切片器索引（int）</param>
    /// <returns>对应的切片器对象</returns>
    IExcelSlicer? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的切片器。
    /// </summary>
    /// <param name="name">切片器名称（string）</param>
    /// <returns>对应的切片器对象</returns>
    IExcelSlicer? this[string name] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 创建新的切片器并返回 <see cref="IExcelSlicer"/> 对象。
    /// </summary>
    /// <param name="slicerDestination">一个字符串，指定工作表的名称，或 Worksheet 表示工作表的对象，将放置生成的切片器。 目标工作表必须位于包含 Slicers 表达式指定的 对象的工作簿中。</param>
    /// <param name="level">如果是 OLAP 数据源，则为创建切片器所基于的级别的序号或多维表达式 (MDX) 名称。 非 OLAP 数据源不支持此参数。</param>
    /// <param name="name">切片器的名称。 如果未指定，Excel 会自动生成一个名称。 </param>
    /// <param name="caption">切片器的标题。</param>
    /// <param name="top">切片器相对于工作表上单元格 A1 左上角的初始垂直位置（以磅为单位）。</param>
    /// <param name="left">切片器相对于工作表上单元格 A1 左上角的初始水平位置（以磅为单位）。</param>
    /// <param name="width">切片器控件的初始宽度（以磅为单位）。</param>
    /// <param name="height">切片器控件的初始高度（以磅为单位）。</param>
    /// <returns></returns>
    IExcelSlicer? Add(string slicerDestination, string? level = null, string? name = null, string? caption = null,
        double? top = null, double? left = null, double? width = null, double? height = null);
}