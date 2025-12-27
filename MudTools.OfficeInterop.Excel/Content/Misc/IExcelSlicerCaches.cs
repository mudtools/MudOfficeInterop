//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作簿中所有切片器缓存的集合，支持遍历、索引和名称访问。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSlicerCaches : IOfficeObject<IExcelSlicerCaches>, IEnumerable<IExcelSlicerCache?>, IDisposable
{
    /// <summary>
    /// 获取集合中切片器缓存的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的切片器缓存。
    /// </summary>
    /// <param name="index">缓存索引</param>
    /// <returns>对应的切片器缓存对象</returns>
    IExcelSlicerCache? this[int index] { get; }

    /// <summary>
    /// 通过名称获取指定的切片器缓存。
    /// </summary>
    /// <param name="name">缓存名称</param>
    /// <returns>对应的切片器缓存对象</returns>
    IExcelSlicerCache? this[string name] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 Workbook）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 创建一个新的切片器缓存并添加到集合中。
    /// </summary>
    /// <param name="source">数据源（PivotTable 或 ListObject 名称）</param>
    /// <param name="field">字段名称（透视表字段或表格列名）</param>
    /// <param name="name">缓存名称（可选，如不提供则自动生成）</param>
    /// <returns>新创建的切片器缓存对象</returns>
    IExcelSlicerCache? Add(string source, string field, string? name = null);

    /// <summary>
    /// 创建一个新的切片器缓存并添加到集合中，支持指定切片器缓存类型。
    /// </summary>
    /// <param name="source">数据源（PivotTable 或 ListObject 名称）</param>
    /// <param name="sourceField">源字段名称（透视表字段或表格列名）</param>
    /// <param name="name">缓存名称（可选，如不提供则自动生成）</param>
    /// <param name="slicerCacheType">切片器缓存类型（可选，默认为标准切片器）</param>
    /// <returns>新创建的切片器缓存对象</returns>
    IExcelSlicerCache? Add2(string source, string sourceField, string? name = null, XlSlicerCacheType? slicerCacheType = null);
}