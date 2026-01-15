//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示切片器缓存中所有层级的集合，仅在 OLAP 数据源中有效。
/// 支持遍历和索引访问。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSlicerCacheLevels : IOfficeObject<IExcelSlicerCacheLevels, MsExcel.SlicerCacheLevels>, IEnumerable<IExcelSlicerCacheLevel?>, IDisposable
{
    /// <summary>
    /// 获取集合中层级的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引（从 1 开始）获取指定的层级。
    /// </summary>
    /// <param name="index">层级索引（1-based）</param>
    /// <returns>对应的层级对象</returns>
    IExcelSlicerCacheLevel? this[int index] { get; }

    /// <summary>
    /// 获取此集合所属的父对象（通常是 SlicerCache）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此集合所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
}