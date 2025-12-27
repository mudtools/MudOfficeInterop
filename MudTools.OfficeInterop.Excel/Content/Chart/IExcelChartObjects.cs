//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// Excel ChartObjects 集合对象的二次封装接口
/// 提供对 Microsoft.Office.Interop.Excel.ChartObjects 的安全访问和操作
/// </summary>
public interface IExcelChartObjects : IExcelComGraphObjects, IEnumerable<IExcelChartObject>, IDisposable
{
    #region 基础属性
    /// <summary>
    /// 获取或设置一个布尔值，指示图表对象是否受到保护
    /// </summary>
    bool ProtectChartObject { get; set; }

    /// <summary>
    /// 获取指定索引的图表对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">图表对象索引（从1开始）</param>
    /// <returns>图表对象</returns>
    IExcelChartObject this[int index] { get; }

    /// <summary>
    /// 获取指定名称的图表对象
    /// </summary>
    /// <param name="name">图表对象名称</param>
    /// <returns>图表对象</returns>
    IExcelChartObject this[string name] { get; }

    /// <summary>
    /// 获取图表对象集合所在的父对象（通常是工作表）
    /// 对应 ChartObjects.Parent 属性
    /// </summary>
    object? Parent { get; }

    #endregion

    #region 创建和添加

    /// <summary>
    /// 向工作表添加新的图表对象
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <returns>新创建的图表对象</returns>
    IExcelChartObject? Add(double left, double top, double width, double height);
    #endregion

    #region 查找和筛选

    /// <summary>
    /// 根据名称查找图表对象
    /// </summary>
    /// <param name="name">图表对象名称</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindByName(string name);

    /// <summary>
    /// 根据位置查找图表对象
    /// </summary>
    /// <param name="left">左边距</param>
    /// <param name="top">顶边距</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindByPosition(double left, double top, double tolerance = 10);

    /// <summary>
    /// 根据大小查找图表对象
    /// </summary>
    /// <param name="width">宽度</param>
    /// <param name="height">高度</param>
    /// <param name="tolerance">容差</param>
    /// <returns>匹配的图表对象数组</returns>
    IExcelChartObject[] FindBySize(double width, double height, double tolerance = 10);

    /// <summary>
    /// 获取指定区域内的所有图表对象
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <returns>区域内的图表对象数组</returns>
    IExcelChartObject[] GetChartsInRange(IExcelRange range);

    /// <summary>
    /// 获取可见的图表对象
    /// </summary>
    /// <returns>可见图表对象数组</returns>
    IExcelChartObject[] GetVisibleCharts();

    #endregion

    #region 操作方法

    /// <summary>
    /// 删除所有图表对象
    /// 对应 ChartObjects.Delete 方法
    /// </summary>
    void Clear();

    /// <summary>
    /// 删除指定索引的图表对象
    /// </summary>
    /// <param name="index">要删除的图表对象索引</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定的图表对象
    /// </summary>
    /// <param name="chartObject">要删除的图表对象</param>
    void Delete(IExcelChartObject chartObject);

    /// <summary>
    /// 批量删除图表对象
    /// </summary>
    /// <param name="indices">要删除的图表对象索引数组</param>
    void DeleteRange(int[] indices);
    #endregion

}