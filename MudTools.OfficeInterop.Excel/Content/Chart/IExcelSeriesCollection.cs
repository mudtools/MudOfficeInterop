//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel SeriesCollection 对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.SeriesCollection 的安全访问和操作
/// </summary>
public interface IExcelSeriesCollection : IEnumerable<IExcelSeries>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取系列集合中的系列数量
    /// 对应 SeriesCollection.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的系列对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">系列索引（从1开始）</param>
    /// <returns>系列对象</returns>
    IExcelSeries this[int index] { get; }

    /// <summary>
    /// 获取系列集合所在的父对象（通常是 Chart）
    /// 对应 SeriesCollection.Parent 属性
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取系列集合所在的 Application 对象
    /// 对应 SeriesCollection.Application 属性
    /// </summary>
    IExcelApplication Application { get; }
    #endregion

    #region 创建和添加
    /// <summary>
    /// 向集合中添加新的空数据系列
    /// 对应 SeriesCollection.NewSeries 方法
    /// </summary>
    /// <returns>新创建的系列对象</returns>
    IExcelSeries Add();

    /// <summary>
    /// 基于数据源创建新的数据系列
    /// 对应 SeriesCollection.Add 方法
    /// </summary>
    /// <param name="source">数据源，可以是 Range、Workbook.WorksheetFunction 或公式字符串</param>
    /// <param name="rowcol">指定数据在源中的排列方式 (1=列, 2=行)</param>
    /// <param name="seriesLabels">是否包含系列标签</param>
    /// <param name="categoryLabels">是否包含分类标签</param>
    /// <returns>新创建的系列对象</returns>
    IExcelSeries CreateSeries(IExcelRange source, int rowcol = 1, bool seriesLabels = false, bool categoryLabels = false);
    #endregion

    #region 操作方法
    /// <summary>
    /// 删除指定索引的系列
    /// </summary>
    /// <param name="index">要删除的系列索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的系列对象
    /// </summary>
    /// <param name="series">要删除的系列对象</param>
    void Delete(IExcelSeries series);

    #endregion
}
