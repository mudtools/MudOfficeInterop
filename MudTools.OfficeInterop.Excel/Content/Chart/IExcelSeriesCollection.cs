//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel SeriesCollection 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.SeriesCollection 的安全访问和操作
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel"), ItemIndex]
public interface IExcelSeriesCollection : IOfficeObject<IExcelSeriesCollection, MsExcel.SeriesCollection>, IEnumerable<IExcelSeries>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取系列集合中的系列数量
    /// 对应 SeriesCollection.Count 属性
    /// </summary>
    int Count { get; }


    /// <summary>
    /// 通过索引获取系列集合中的元素
    /// </summary>
    /// <param name="index">要获取的系列的从零开始的索引</param>
    /// <returns>指定索引处的 Excel 系列对象</returns>
    IExcelSeries this[int index] { get; }

    /// <summary>
    /// 通过名称获取系列集合中的元素
    /// </summary>
    /// <param name="name">要获取的系列的名称</param>
    /// <returns>具有指定名称的 Excel 系列对象</returns>
    IExcelSeries this[string name] { get; }

    /// <summary>
    /// 获取系列集合所在的父对象（通常是 Chart）
    /// 对应 SeriesCollection.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取系列集合所在的 Application 对象
    /// 对应 SeriesCollection.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }
    #endregion

    /// <summary>
    /// 创建一个新的数据系列并将其添加到集合中
    /// </summary>
    /// <returns>新创建的数据系列对象，如果创建失败则返回null</returns>
    IExcelSeries? NewSeries();

    /// <summary>
    /// 扩展现有数据系列集合，将新的数据源添加到现有系列中
    /// </summary>
    /// <param name="source">数据源，可以是Range对象或其他数据源</param>
    /// <param name="rowcol">指定数据排列方式，按行(XlRowCol.xlRows)或按列(XlRowCol.xlColumns)</param>
    /// <param name="categoryLabels">是否将第一行/列作为分类标签处理</param>
    /// <returns>扩展操作的结果对象</returns>
    [IgnoreGenerator]
    object Extend(object source, XlRowCol rowcol = XlRowCol.xlRows, bool? categoryLabels = null);

    /// <summary>
    /// 向集合中添加新的数据系列
    /// </summary>
    /// <param name="source">数据源，可以是Range对象或其他数据源</param>
    /// <param name="rowcol">指定数据排列方式，按行(XlRowCol.xlRows)或按列(XlRowCol.xlColumns)</param>
    /// <param name="seriesLabels">是否将第一行/列作为系列标签处理</param>
    /// <param name="categoryLabels">是否将第一行/列作为分类标签处理</param>
    /// <param name="replace">是否替换现有的冲突系列</param>
    /// <returns>新添加的数据系列对象，如果添加失败则返回null</returns>
    [IgnoreGenerator]
    IExcelSeries? Add(object source,
        XlRowCol rowcol = XlRowCol.xlRows,
        bool? seriesLabels = null,
        bool? categoryLabels = null,
        bool? replace = null);

    /// <summary>
    /// 将剪贴板中的数据粘贴到系列集合中
    /// </summary>
    /// <param name="rowcol">指定数据排列方式，按行(XlRowCol.xlRows)或按列(XlRowCol.xlColumns)</param>
    /// <param name="seriesLabels">是否将第一行/列作为系列标签处理</param>
    /// <param name="categoryLabels">是否将第一行/列作为分类标签处理</param>
    /// <param name="replace">是否替换现有的冲突系列</param>
    /// <param name="newSeries">是否创建新系列</param>
    /// <returns>粘贴操作的结果对象</returns>
    object Paste(XlRowCol rowcol = XlRowCol.xlRows, bool? seriesLabels = null,
         bool? categoryLabels = null, bool? replace = null, bool? newSeries = null);
}
