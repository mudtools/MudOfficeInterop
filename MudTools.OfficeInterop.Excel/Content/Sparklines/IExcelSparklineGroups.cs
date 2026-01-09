//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示迷你图组的集合。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface IExcelSparklineGroups : IEnumerable<IExcelSparklineGroup?>, IOfficeObject<IExcelSparklineGroups, MsExcel.SparklineGroups>, IDisposable
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
    /// 创建新的迷你图组并返回 SparklineGroup 对象。
    /// </summary>
    /// <param name="type">迷你图类型。</param>
    /// <param name="sourceData">用于创建迷你图的源数据区域。</param>
    /// <returns>新创建的迷你图组对象。</returns>
    IExcelSparklineGroup? Add(XlSparkType type, string sourceData);

    /// <summary>
    /// 获取关联 Range 对象中的迷你图组数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的 SparklineGroup 对象。
    /// </summary>
    /// <param name="index">集合中元素的位置索引。</param>
    /// <returns>指定的迷你图组对象。</returns>
    IExcelSparklineGroup? this[int index] { get; }

    /// <summary>
    /// 通过索引获取集合中的 SparklineGroup 对象。
    /// </summary>
    /// <param name="name">集合中元素的位置索引。</param>
    /// <returns>指定的迷你图组对象。</returns>
    IExcelSparklineGroup? this[string name] { get; }

    /// <summary>
    /// 清除选定的迷你图。
    /// </summary>
    void Clear();

    /// <summary>
    /// 清除选定的迷你图组。
    /// </summary>
    void ClearGroups();

    /// <summary>
    /// 将选定的迷你图分组。
    /// </summary>
    /// <param name="location">组中第一个单元格的位置。</param>
    void Group(IExcelRange location);

    /// <summary>
    /// 取消选定迷你图组的分组。
    /// </summary>
    void Ungroup();
}